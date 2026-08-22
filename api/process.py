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
            # 1. リクエストボディの解析
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body.decode('utf-8'))

            file_b64 = data.get("file")
            mime_type = data.get("mimeType", "image/png")

            if not file_b64:
                self.send_error(400, "No file provided.")
                return

            # 2. Gemini APIの設定
            api_key = os.environ.get("GEMINI_API_KEY")
            if not api_key:
                self.send_error(500, "GEMINI_API_KEY is not set.")
                return

            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash')

            # 3. Geminiによる画像解析・ルート計算の実行プロンプト
            prompt = """
添付された画像/PDFから全ての所在地・店舗・施設名・駅名を読み取り、以下の厳格なルールに基づいて解析を行ってください。

【最優先処理ルール】
1. 画像内の略称・手書き文字・通称は、以下の【正式名称・最新情報】に必ず自動変換してください。
   - みらい千林西 ➔ 関西みらい銀行 千林西支店
   - 関西みらい 出来島 ➔ 関西みらい銀行 出来島支店
   - フロンティア 南大阪 ➔ フロンティア不動産販売 南大阪店
   - 池田泉州銀行 交野支店 ➔ 池田泉州銀行 交野支店（最寄り：JR星田駅）
   - ろうきん / 近畿労金 ➔ 近畿労働金庫
   - だいしん / 大阪信金 ➔ 大阪信用金庫
   - 京信 ➔ 京都信用金庫
   - 中信 ➔ 京都中央信用金庫
   - あましん ➔ 尼崎信用金庫
   - ひまわり ➔ 兵庫ひまわり信用組合
   - ひょうしん ➔ 兵庫信用金庫
   - 京滋 ➔ 京滋信用組合
   - 京銀 ➔ 京都銀行
   - ひらしん ➔ 枚方信用金庫
   - ドリーム ➔ ドリームホーム
   - ドリーム本社 ➔ dreamtown本社
   - ドリーム洛西口 ➔ ドリームホーム 洛西口駅前店

2. 各拠点間の「明日の10:30出発」における移動データを算出して、以下の純粋なJSON形式のみで出力してください（Markdownの ```json 等の囲みは不要です）。

【JSON出力フォーマット】
{
  "locations": [
    {
      "no": 1,
      "raw_name": "画像上の表記",
      "official_name": "特定された正式店舗・施設名",
      "address": "正式住所",
      "stations": "最寄り駅一覧と徒歩/タクシー所要時間"
    }
  ],
  "transit": [
    ["出発地/目的地", "拠点1の正式名称", "拠点2の正式名称"],
    ["拠点1の正式名称", "同地点", "所要時間・経由・乗換"],
    ["拠点2の正式名称", "所要時間・経由・乗換", "同地点"]
  ],
  "driving": [
    ["出発地/目的地", "拠点1の正式名称", "拠点2の正式名称"],
    ["拠点1の正式名称", "同地点", "〇分 (〇km)"],
    ["拠点2の正式名称", "〇分 (〇km)", "同地点"]
  ],
  "walking": [
    ["出発地/目的地", "拠点1の正式名称", "拠点2の正式名称"],
    ["拠点1の正式名称", "同地点", "〇分 (〇km)"],
    ["拠点2の正式名称", "〇分 (〇km)", "同地点"]
  ]
}
"""

            image_bytes = base64.b64decode(file_b64)
            image_part = {"mime_type": mime_type, "data": image_bytes}
            
            response = model.generate_content([prompt, image_part])
            res_text = response.text.strip()
            if res_text.startswith("```"):
                res_text = res_text.split("\n", 1)[1].rsplit("```", 1)[0].strip()

            result_data = json.loads(res_text)

            # 4. Excelブックの組み立て
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

            # マトリックスシートの作成
            modes = [("公共交通機関", "transit", "transit"), ("車利用", "driving", "driving"), ("徒歩利用", "walking", "walking")]
            
            for idx, (title, data_key, mode_param) in enumerate(modes):
                ws = wb.active if idx == 0 else wb.create_sheet(title=title)
                if idx == 0:
                    ws.title = title
                
                ws.freeze_panes = 'B5'
                ws['A1'] = "出発想定日時：明日の 10:30 出発"
                ws['A1'].font = Font(bold=True, size=12)
                
                ws['A4'] = "出発地 ＼ 目的地"
                ws['A4'].fill = header_fill
                ws['A4'].font = header_font
                ws['A4'].alignment = center_align

                matrix_data = result_data.get(data_key, [])
                if matrix_data:
                    # ヘッダー設置
                    for c_idx, col_name in enumerate(matrix_data[0][1:], start=2):
                        cell = ws.cell(row=4, column=c_idx, value=col_name)
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = center_align

                    # データ設置 & Google Maps ハイパーリンク設定
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
                                # Google Maps リンク付与
                                maps_url = f"[https://www.google.com/maps/dir/?api=1&origin=](https://www.google.com/maps/dir/?api=1&origin=){urllib.parse.quote(origin_name)}&destination={urllib.parse.quote(dest_name)}&travelmode={mode_param}"
                                cell.hyperlink = maps_url
                                cell.font = Font(color="0000FF", underline="single")

            # 全体マップシートの作成
            ws_map = wb.create_sheet(title="全体マップ")
            ws_map.append(["No.", "画像上の表記", "特定された正式店舗・施設名", "正式住所", "Googleマップリンク", "利用可能な最寄り駅一覧と所要時間"])
            for cell in ws_map[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align

            for loc in result_data.get("locations", []):
                official = loc.get("official_name", "")
                map_url = f"[https://www.google.com/maps/search/?api=1&query=](https://www.google.com/maps/search/?api=1&query=){urllib.parse.quote(official)}"
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

            # 5. バイナリデータとしてダウンロード出力
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_data = excel_buffer.getvalue()

            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', 'attachment; filename="matrix.xlsx"')
            self.end_headers()
            self.wfile.write(excel_data)

        except Exception as e:
            self.send_error(500, str(e))
