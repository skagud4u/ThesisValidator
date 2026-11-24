import os
import re
import json
import subprocess
import tempfile
from difflib import SequenceMatcher
import pandas as pd
import pdfplumber
import pymupdf
from openpyxl.styles import PatternFill, Font, Alignment

class ThesisValidator:
    def __init__(self, target_dir="./Sample", output_excel="검수결과_리포트.xlsx"):
        self.target_dir = target_dir
        self.output_excel = output_excel

    def _ocr_page_with_windows(self, file_path, page_index):
        """이미지 기반 PDF 페이지를 렌더링하여 Windows 내장 한국어 OCR로 텍스트 추출"""
        tmp_png = None
        try:
            with pymupdf.open(file_path) as doc:
                pix = doc[page_index].get_pixmap(dpi=200)
                fd, tmp_png = tempfile.mkstemp(suffix=".png")
                os.close(fd)
                pix.save(tmp_png)

            ps_script = f"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime]

function Await($WinRtTask, $ResultType) {{
    $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {{
        $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    }})[0].MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    return $netTask.Result
}}

$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage((New-Object Windows.Globalization.Language("ko")))
if ($null -eq $engine) {{ exit 1 }}

$fs = [System.IO.File]::OpenRead("{tmp_png.replace(chr(92), '/')}")
$stream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($fs)
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
$fs.Dispose()

$items = @()
foreach ($line in $result.Lines) {{
    $minY = ($line.Words | ForEach-Object {{ $_.BoundingRect.Y }} | Measure-Object -Minimum).Minimum
    $minX = ($line.Words | ForEach-Object {{ $_.BoundingRect.X }} | Measure-Object -Minimum).Minimum
    $items += [PSCustomObject]@{{ y = $minY; x = $minX; text = $line.Text }}
}}
$items | ConvertTo-Json -Compress
"""
            proc = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                capture_output=True,
                text=True,
                encoding="utf-8",
                timeout=30
            )
            if proc.returncode != 0 or not proc.stdout.strip():
                return ""

            data = json.loads(proc.stdout.strip())
            if isinstance(data, dict):
                data = [data]

            # Y좌표 기준 정렬 후 같은 행(15px 이내)끼리 X좌표 순으로 병합
            data.sort(key=lambda item: (item["y"], item["x"]))
            rows = []
            for item in data:
                if rows and abs(item["y"] - rows[-1][0]["y"]) <= 15:
                    rows[-1].append(item)
                else:
                    rows.append([item])

            lines = []
            for row in rows:
                row.sort(key=lambda item: item["x"])
                lines.append(" ".join(item["text"] for item in row))
            return "\n".join(lines)
        except Exception:
            return ""
        finally:
            if tmp_png and os.path.exists(tmp_png):
                try:
                    os.remove(tmp_png)
                except OSError:
                    pass
        
    def check_thesis(self, file_path):
        filename = os.path.basename(file_path)
        result = {
            "파일명": filename,
            "검수결과": "통과",
            "구성순서_오류": "정상",
            "인준서_지도교수위치": "정상",
            "페이지번호_오류": "정상",
            "상세_반려사유": [],
            "학생_안내문구": ""
        }
        
        try:
            with pdfplumber.open(file_path) as pdf:
                total_pages = len(pdf.pages)
                pages_text = []
                pages_footer = []

                for idx, page in enumerate(pdf.pages):
                    text = page.extract_text() or ""
                    # 서두부(1~10p)에 텍스트가 없고 이미지로만 된 페이지(인준서 스캔본 등)는 OCR 수행
                    if idx < 10 and len(text.strip()) < 20 and len(page.images) > 0:
                        ocr_text = self._ocr_page_with_windows(file_path, idx)
                        if ocr_text:
                            text = ocr_text
                    pages_text.append((idx + 1, text))
                    
                    # 하단 10% 영역의 페이지 번호 텍스트 추출
                    h = page.height
                    w = page.width
                    footer_crop = page.crop((0, h * 0.90, w, h))
                    footer_text = footer_crop.extract_text() or ""
                    footer_num = re.sub(r'[^a-zA-Z0-9]', '', footer_text.strip())
                    pages_footer.append((idx + 1, footer_num.lower()))

                # 1. 필수 구성 항목 위치 식별
                sections = {
                    "인준서": None,
                    "국문초록": None,
                    "목차": None,
                    "본문_시작": None,
                    "참고문헌": None,
                    "영문초록": None
                }

                for p_num, text in pages_text:
                    clean_t = text.replace(" ", "")
                    if not sections["인준서"] and ("인준" in clean_t and ("심사위원" in clean_t or "위원장" in clean_t or "지도교수" in clean_t or "논문으로인준" in clean_t)):
                        sections["인준서"] = p_num
                    elif not sections["국문초록"] and ("국문초록" in clean_t or "초록" in clean_t):
                        sections["국문초록"] = p_num
                    elif not sections["목차"] and "목차" in clean_t:
                        sections["목차"] = p_num
                    elif not sections["본문_시작"] and ("제1장" in clean_t or "I.서론" in clean_t):
                        sections["본문_시작"] = p_num
                    elif not sections["참고문헌"] and "참고문헌" in clean_t:
                        sections["참고문헌"] = p_num
                    elif not sections["영문초록"] and ("ABSTRACT" in clean_t.upper()):
                        sections["영문초록"] = p_num

                # 2. 순서 유효성 검증
                # 기대 순서: 인준서 < 국문초록 < 목차 < 참고문헌
                order_keys = ["인준서", "국문초록", "목차", "참고문헌"]
                prev_page = 0
                for k in order_keys:
                    p = sections.get(k)
                    if p is None:
                        result["구성순서_오류"] = "항목누락"
                        result["상세_반려사유"].append(f"필수 구성 항목인 [{k}] 미검출")
                    elif p < prev_page:
                        result["구성순서_오류"] = "순서역전"
                        result["상세_반려사유"].append(f"구성 순서 오류: [{k}]({p}p)가 이전 항목보다 앞섬")
                    else:
                        prev_page = p

                # 3. 인준서 지도교수 배치 확인 (최하단 독립 배치 및 심사위원 내 배치 모두 정상 처리)
                if sections["인준서"]:
                    result["인준서_지도교수위치"] = "정상"

                # 4. 페이지 번호 체계 검증
                # 초록~목차 구간: 로마자 소문자 검사
                start_p = sections["국문초록"] or 3
                end_p = sections["본문_시작"] or (start_p + 2)
                for p_num, f_num in pages_footer[start_p - 1: end_p - 1]:
                    if f_num and not re.match(r'^[ivxlcdm]+$', f_num):
                        if re.match(r'^\d+$', f_num):
                            result["페이지번호_오류"] = "로마자누락"
                            result["상세_반려사유"].append(f"{p_num}페이지 서두부에 아라비아 숫자 사용 (로마자 표기 필요)")
                            break

        except Exception as e:
            result["검수결과"] = "검사오류"
            result["상세_반려사유"].append(f"파일 판독 에러: {str(e)}")

        # 최종 판정 및 안내문구 합성
        if result["상세_반려사유"]:
            result["검수결과"] = "반려"
            reasons = "\n".join([f"- {r}" for r in result["상세_반려사유"]])
            result["학생_안내문구"] = f"[학위논문 체제 보완 요청]\n제출하신 논문에서 아래 규정 위반 항목이 발견되어 반려 처리됩니다.\n{reasons}\n체제 지침을 확인 후 재제출 바랍니다."
        else:
            result["학생_안내문구"] = "모든 필수 체제 규격 통과"

        result["상세_반려사유"] = " | ".join(result["상세_반려사유"])
        return result

    def run_batch_and_export(self):
        records = []
        if not os.path.exists(self.target_dir):
            print(f"디렉토리를 찾을 수 없습니다: {self.target_dir}")
            return

        files = [f for f in os.listdir(self.target_dir) if f.lower().endswith(".pdf")]
        print(f"총 {len(files)}개의 논문 파일 검수를 시작합니다...")

        for f in files:
            full_path = os.path.join(self.target_dir, f)
            res = self.check_thesis(full_path)
            records.append(res)

        df = pd.DataFrame(records)
        df.to_excel(self.output_excel, index=False)

        # 엑셀 서식 적용 (반려 셀 붉은색 하이라이트)
        wb = pd.ExcelFile(self.output_excel)
        from openpyxl import load_workbook
        wb = load_workbook(self.output_excel)
        ws = wb.active

        red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

        for row in range(2, len(records) + 2):
            cell = ws[f"B{row}"]  # 검수결과 컬럼
            if cell.value == "반려":
                cell.fill = red_fill
            elif cell.value == "통과":
                cell.fill = green_fill
            ws[f"G{row}"].alignment = Alignment(wrap_text=True)

        wb.save(self.output_excel)
        print(f"검수 완료! 결과 파일 저장: {self.output_excel}")

if __name__ == "__main__":
    validator = ThesisValidator(target_dir="./Sample")
    validator.run_batch_and_export()
