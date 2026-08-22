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
            # 最新の安定型モデルに修正
            model = genai.GenerativeModel('gemini-2.0-flash')

            prompt = f"""
添付資料（画像またはPDF）に記載されている全ての施設名・店舗名・住所・駅名を読み取り、以下の【略称変換ルール】を厳格に適用して正式名称を特定し、「{dept_time}」における移動所要時間マトリックスデータをJSONのみで作成してください。

【略称変換ルール】
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

【出力JSON構造】
```json
{{
  "locations": [
    {{ "no": 1, "raw_name": "画像表記", "official_name": "正式名称", "address": "住所", "stations": "最寄り駅一覧" }}
  ],
  "transit": [
    ["出発地／目的地", "拠点A", "拠点B"],
    ["拠点A", "同地点", "所要時間・乗換"],
    ["拠点B", "所要時間・乗換", "同地点"]
  ],
  "driving": [
    ["出発地／目的地", "拠点A", "拠点B"],
    ["拠点A", "同地点", "〇分 (〇km)"],
    ["拠点B", "〇分 (〇km)", "同地点"]
  ],
  "walking": [
    ["出発地／目的地", "拠点A", "拠点B"],
    ["拠点A", "同地点", "〇分 (〇km)"],
    ["拠点B", "〇分 (〇km)", "同地点"]
  ]
}}
