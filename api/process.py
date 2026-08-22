import os
import io
import json
import base64
import urllib.parse
from http.server import BaseHTTPRequestHandler
import google.generativeai as genai
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body.decode('utf-8'))

            file_b64 = data.get("file")
            mime_type = data.get("mimeType", "image/png")
            dept_time = data.get("deptTime", "明日の 10:30 出発")

            if not file_b64:
                self.send_error(400, "ファイルデータが送信されていません。")
                return

            api_key = os.environ.get("GEMINI_API_KEY")
            if not api_key:
                self.send_error(500, "Vercelの環境変数 GEMINI_API_KEY が設定されていません。")
                return

            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-2.0-flash')

            prompt = (
                "添付資料（画像またはPDF）に記載されている全ての施設名・店舗名・住所・駅名を読み取り、"
                "以下の【略称変換ルール】を厳格に適用して正式名称を特定し、"
                f"「{dept_time}」における移動所要時間マトリックスデータをJSONのみで作成してください。\n\n"
                "【略称変換ルール】\n"
                "- みらい千林西 ➔ 関西みらい銀行 千林西支店\n"
                "- 関西みらい 出来島 ➔ 関西みらい銀行 出来島支店\n"
                "- フロンティア 南大阪 ➔ フロンティア不動産販売 南大阪店\n"
                "- 池田泉州銀行 交野支店 ➔ 池田泉州銀行 交野支店（最寄り：JR星田駅）\n"
                "- ろうきん / 近畿労金 ➔ 近畿労働金庫\n"
                "- だいしん / 大阪信金 ➔ 大阪信用金庫\n"
                "- 京信 ➔ 京都信用金庫\n"
                "- 中信 ➔ 京都中央信用金庫\n"
                "- あましん ➔ 尼崎信用金庫\n"
                "- ひまわり ➔ 兵庫ひまわり信用組合\n"
                "- ひょうしん ➔ 兵庫信用金庫\n"
                "- 京滋 ➔ 京滋信用組合\n"
                "- 京銀 ➔ 京都銀行\n"
                "- ひらしん ➔ 枚方信用金庫\n"
                "- ドリーム ➔ ドリームホーム\n"
                "- ドリーム本社 ➔ dreamtown本社\n"
                "- ドリーム洛西口 ➔ ドリームホーム 洛西口駅前店\n\n"
                "【出力JSON構造】\n"
                "{\n"
                '  "locations": [\n'
                '    { "no": 1, "raw_name": "画像表記", "official_name": "正式名称", "address": "住所", "stations": "最寄り駅一覧" }\n'
                "  ],\n"
                '  "transit": [\n'
                '    ["出発地／目的地", "拠点A", "拠点B"],\n'
                '    ["拠点A", "同地点", "所要時間・乗換"],\n'
                '    ["拠点B", "所要時間・乗換", "同地点"]\n'
                "  ],\n"
                '  "driving": [\n'
                '    ["出発地／目的地", "拠点A", "拠点B"],\n'
                '    ["拠点A", "同地点", "〇分 (〇km)"],\n'
                '    ["拠点B", "〇分 (〇km)", "同地点"]\n'
                "  ],\n"
                '  "walking": [\n'
                '    ["出発地／目的地", "拠点A", "拠点B"],\n'
                '    ["拠点A", "同地点", "〇分 (〇km)"],\n'
                '    ["拠点B", "〇分 (〇km)", "同地点"]\n'
                "  ]\n"
                "}\n"
            )

            file_bytes = base64.b64decode(file_b64)
            part = {"mime_type": mime_type, "data": file_bytes}
            
            response = model.generate_content([prompt, part])
            res_text = response.text.strip()
            
            if "```json" in res_text:
                res_text = res_text.split("```json")[1].split("```")[0].strip()
            elif "```" in res_text:
                res_text = res_text.split("```")[1].split("```")[0].strip()

            result_data = json.loads(res_text)

            wb = openpyxl.Workbook()
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )
            center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

            modes = [("公共交通機関", "transit", "transit"), ("車利用", "driving", "driving"), ("徒歩利用", "walking", "walking")]
            
            for idx, (title, data_key, mode_param) in enumerate(modes):
                ws = wb.active if idx == 0 else wb.create_sheet(title=title)
                if idx == 0:
                    ws.title = title
                
                ws.freeze_panes = 'B5'
                ws['A1'] = f"出発想定日時：{dept_time}"
                ws['A1'].font = Font(bold=True, size=12)
                
                ws['A4'] = "出発地 ＼ 目的地"
                ws['A4'].fill = header_fill
                ws['A4'].font = header_font
                ws['A4'].alignment = center_align

                matrix_data = result_data.get(data_key, [])
                if matrix_data:
                    for c_idx, col_name in enumerate(matrix_data[0][1:], start=2):
                        cell = ws.cell(row=4, column=c_idx, value=col_name)
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = center_align

                    for r_idx, row_data in enumerate(matrix_data[1:], start=5):
                        origin_name = row_data[0]
                        ws.cell(row=r_idx, column=1, value=origin_name).font = Font(bold=True)
                        ws.cell(row=r_idx, column=1).fill = PatternFill(start_color="F2F2F2", fill_type="solid")

                        for c_idx, val in enumerate(row_data[1:], start=2):
                            dest_name = matrix_data[0][c_idx-1]
                            cell = ws.cell(row=r_idx, column=c_idx)
                            cell.alignment = center_align
                            cell.border = thin_border

                            if origin_name == dest_name or val == "同地点":
                                cell.value = "-"
                            else:
                                cell.value = val
                                maps_url = f"https://www.google.com/maps/dir/?api=1&origin={urllib.parse.quote(origin_name)}&destination={urllib.parse.quote(dest_name)}&travelmode={mode_param}"
                                cell.hyperlink = maps_url
                                cell.font = Font(color="0000FF", underline="single")

            ws_map = wb.create_sheet(title="全体マップ")
            ws_map.append(["No.", "画像上の表記", "特定された正式店舗・施設名", "正式住所", "Googleマップリンク", "利用可能な最寄り駅一覧と所要時間"])
            for cell in ws_map[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align

            for loc in result_data.get("locations", []):
                official = loc.get("official_name", "")
                map_url = f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(official)}"
                row_cells = [
                    loc.get("no"),
                    loc.get("raw_name"),
                    official,
                    loc.get("address"),
                    map_url,
                    loc.get("stations")
                ]
                ws_map.append(row_cells)
                last_row = ws_map.max_row
                ws_map.cell(row=last_row, column=5).hyperlink = map_url
                ws_map.cell(row=last_row, column=5).font = Font(color="0000FF", underline="single")

            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_data = excel_buffer.getvalue()

            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', 'attachment; filename="matrix.xlsx"')
            self.end_headers()
            self.wfile.write(excel_data)

        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'text/plain; charset=utf-8')
            self.end_headers()
            self.wfile.write(f"処理エラーが発生しました:\n{str(e)}".encode('utf-8'))
