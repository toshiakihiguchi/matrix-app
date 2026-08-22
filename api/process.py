import os
import io
import json
import urllib.parse
from http.server import BaseHTTPRequestHandler
import google.generativeai as genai
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- 1. 厳格な名称変換辞書 ---
NAME_DICTIONARY = {
    "みらい千林西": {"name": "関西みらい銀行 千林西支店", "address": "大阪府大阪市旭区千林2-11-24"},
    "関西みらい 出来島": {"name": "関西みらい銀行 出来島支店", "address": "大阪府大阪市西淀川区出来島1-1-2"},
    "フロンティア 南大阪": {"name": "フロンティア不動産販売 南大阪店", "address": "大阪府堺市北区百舌鳥梅町1丁15-2"},
    "池田泉州銀行 交野支店": {"name": "池田泉州銀行 交野支店", "address": "大阪府交野市星田5-12-2", "base_station": "JR星田駅"},
    "ろうきん": {"name": "近畿労働金庫", "address": ""},
    "だいしん": {"name": "大阪信用金庫", "address": ""},
    "京信": {"name": "京都信用金庫", "address": ""},
    "中信": {"name": "京都中央信用金庫", "address": ""},
    "あましん": {"name": "尼崎信用金庫", "address": ""},
    "ひまわり": {"name": "兵庫ひまわり信用組合", "address": ""},
    "ひょうしん": {"name": "兵庫信用金庫", "address": ""},
    "京滋": {"name": "京滋信用組合", "address": ""},
    "京銀": {"name": "京都銀行", "address": ""},
    "ひらしん": {"name": "枚方信用金庫", "address": ""},
    "ドリーム": {"name": "ドリームホーム", "address": ""},
    "ドリーム本社": {"name": "dreamtown本社", "address": "京都府京都市中京区壬生坊城町"},
    "ドリーム洛西口": {"name": "ドリームホーム 洛西口駅前店", "address": "京都府向日市寺戸町七ノ坪"}
}

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            # APIキーの取得
            api_key = os.environ.get("GEMINI_API_KEY")
            if not api_key:
                self.send_error(500, "GEMINI_API_KEY is not set.")
                return

            genai.configure(api_key=api_key)

            # --- Excelブック作成 ---
            wb = openpyxl.Workbook()
            
            # スタイル定義（紺色ヘッダー・格子枠・白文字）
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )
            center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

            # 4シート構成
            sheets = {
                "transit": wb.active,
                "driving": wb.create_sheet(title="車利用"),
                "walking": wb.create_sheet(title="徒歩利用"),
                "map": wb.create_sheet(title="全体マップ")
            }
            sheets["transit"].title = "公共交通機関"

            # 1~3. マトリックスシートの基本構築とウィンドウ枠固定（B5基準）
            modes = [("transit", "transit"), ("driving", "driving"), ("walking", "walking")]
            for sheet_key, mode in modes:
                ws = sheets[sheet_key]
                ws.freeze_panes = 'B5'  # 要求通りB5セル基準でスクロール固定
                
                # ヘッダー設定
                ws['A1'] = "出発想定日時：明日の 10:30 出発"
                ws['A1'].font = Font(bold=True, size=12)
                ws['A4'] = "出発地 ＼ 目的地"
                ws['A4'].fill = header_fill
                ws['A4'].font = header_font
                ws['A4'].alignment = center_align

            # 4. 全体マップシート構築
            ws_map = sheets["map"]
            ws_map.append(["No.", "画像上の表記", "特定された正式店舗・施設名", "正式住所", "Googleマップリンク", "利用可能な最寄り駅一覧と所要時間"])
            for cell in ws_map[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align

            # バイナリ出力処理
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
