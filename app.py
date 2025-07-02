from flask import Flask, request, send_file, render_template_string 
import pyperclip
from icalendar import Calendar,Event, vText
import datetime
import re
import os
from dateutil import parser
import io

app= Flask(__name__)

def parse_event_from_text(text_content):
    """
    テキストコンテンツからイベント情報を解析します。
    この部分は、ユーザーがコピーするテキストの形式に合わせて調整が必要です。
    """
    title = "新しいカレンダーイベント"
    start_dt = datetime.datetime.now()
    description = text_content
    location = ""

    # 1. 日付の解析 (dateutil.parser を使用すると柔軟性が高い)
    try:
        # 例: "2025年7月1日", "来週月曜日", "明日", "7/1" など
        # 'fuzzy=True' で、日付以外の文字列も含むテキストから日付を抽出試行
        parsed_date = parser.parse(text_content, fuzzy=True, dayfirst=False, yearfirst=True)
        start_dt = start_dt.replace(year=parsed_date.year, month=parsed_date.month, day=parsed_date.day)
    except parser.ParserError:
        pass # 日付が見つからない場合は現在の日のまま

    # 2. 時刻の解析 (日付が見つかった場合、その日の時刻として解析)
    time_match = re.search(r'(\d{1,2}:\d{2}(?:[APap][Mm])?|\d{1,2}時\d{2}分?|\d{1,2}時)', text_content)
    if time_match:
        try:
            time_str = time_match.group(1).replace('時', ':').replace('分', '')
            # AM/PMを含む場合の処理
            if 'am' in time_str.lower() or 'pm' in time_str.lower():
                parsed_time = datetime.datetime.strptime(time_str, '%I:%M%p').time()
            elif ':' in time_str:
                parsed_time = datetime.datetime.strptime(time_str, '%H:%M').time()
            else: # 例: "10時" の場合
                parsed_time = datetime.datetime.strptime(time_str + ':00', '%H:%M').time()

            start_dt = start_dt.replace(hour=parsed_time.hour, minute=parsed_time.minute, second=0, microsecond=0)

            # もし解析された時刻が現在時刻より過去で、かつ日付が今日の場合、翌日にする（午前/午後が不明な場合など）
            if start_dt < datetime.datetime.now() and start_dt.date() == datetime.datetime.now().date():
                start_dt += datetime.timedelta(days=1)

        except ValueError:
            pass # 解析失敗時はデフォルト値を使用

    # 3. タイトルの解析 (最も重要な情報)
    # 例: "〇〇と打ち合わせ", "〇〇会議", "〇〇の予約" など
    # ここは非常にカスタマイズが必要です。
    # 簡単な例として、日付や時刻の前後にあるキーワードをタイトルとする
    # あるいは、特定のキーワード（例: "会議", "打ち合わせ", "予約"）を見つけて、その前後のテキストをタイトルとする
    title_keywords = ["会議", "打ち合わせ", "ミーティング", "予約", "イベント", "リマインダー"]
    found_title = False
    for keyword in title_keywords:
        if keyword in text_content:
            # キーワードの前の部分をタイトルにする（簡易的）
            parts = text_content.split(keyword, 1)
            if parts[0].strip():
                title = parts[0].strip() + keyword
            else: # キーワードが先頭に近い場合など
                title = text_content.replace('\n', ' ')[:100] # 長すぎないように
            found_title = True
            break
    
    if not found_title:
        # 日付や時刻の文字列を除去して残りをタイトルにする（より高度な方法）
        cleaned_text = re.sub(r'(\d{4}[年/]\d{1,2}[月/]\d{1,2}日?|\d{1,2}:\d{2}(?:[APap][Mm])?|\d{1,2}時\d{2}分?|\d{1,2}時)', '', text_content)
        cleaned_text = cleaned_text.replace('\n', ' ').strip()
        if cleaned_text:
            title = cleaned_text[:100] + '...' if len(cleaned_text) > 100 else cleaned_text
        else:
            title = "クリップボードからのイベント"

    # 4. 説明 (残りのテキスト全て、または特定の情報)
    description = text_content

    # 5. 場所の解析 (例: "場所：〇〇", "@〇〇" など)
    location_match = re.search(r'(場所|@|にて)[:：]?\s*(.+?)(?=\n|$)', text_content)
    if location_match:
        location = location_match.group(2).strip()

    return title, start_dt, description, location

#webアプリのルートを定義

#メインページ（入力フォーム)
@app.route('/',methods=['GET'])
def index():
    #ユーザーがテキスト入力するHTMLフォーム
    return render_template_string("""
        <!DOCTYPE html>
        <html lang="ja">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>カレンダー自動登録</title>
            <style>
                body { font-family: sans-serif; margin: 20px; text-align: center; }
                textarea { width: 90%; max-width: 500px; height: 150px; margin-bottom: 10px; padding: 10px; font-size: 16px; border: 1px solid #ccc; border-radius: 5px; }
                button { padding: 10px 20px; font-size: 18px; background-color: #007bff; color: white; border: none; border-radius: 5px; cursor: pointer; }
                button:hover { background-color: #0056b3; }
                h1 { color: #333; }
                p { color: #666; }
            </style>
        </head>
        <body>
            <h1>テキストからカレンダーイベントを生成</h1>
            <p>ここにカレンダーに登録したいテキストをペーストしてください。</p>
            <form action="/generate_ical" method="post">
                <textarea name="event_text" id="event_text" placeholder="例: 来週月曜日10:30からオンラインミーティング、議題は新プロジェクトについて。" autofocus></textarea><br>
                <button type="submit">カレンダーイベントを作成</button>
            </form>
            <p><small>（生成されたファイルをタップしてカレンダーにインポートしてください。）</small></p>
        </body>
        </html>                                                  
         """)

#iCalファイルを生成してダウンロードさせるエンドポイント
@app.route('/generate_ical',methods=['POST'])
def generate_ical():
    event_text=request.form['event_text']

    if not event_text:
        return "テキストが入力されてません",400
    
    try:
        title, start_dt, description, location = parse_event_from_text(event_text)

        cal= Calendar()
        cal.add('prodid', '-// My Python Calender Helper Web//example.com//')
        cal.add('version','2.0')

        event=Event()
        event.add('summary', vText(title))
        event.add('dtstart', start_dt)
        event.add('dtend', start_dt + datetime.timedelta(hours=1)) # デフォルトで1時間後
        event.add('description', vText(description))
        if location:
            event.add('location', vText(location))
        event.add('priority', 5)

        cal.add_component(event)

        # iCalデータをメモリ上で生成し、ファイルとして提供
        ical_data = cal.to_ical()
        file_name = f"{title.replace(' ', '_').replace('/', '_')}.ics"
        
        # io.BytesIO を使ってメモリからファイルを送信
        return send_file(
            io.BytesIO(ical_data),
            mimetype='text/calendar',
            as_attachment=True,
            download_name=file_name
        )
    except Exception as e:
        return f"エラーが発生しました: {e}", 500
    
if __name__=='__main__':
    #serverを起動
    #iPhoneからアクセスするためにはhost='0.0.0.0'に設定して
    #ファイアウォールやルータの設定でポート（デフォルト5000）を開くことが必要かも
    app.run(debug=True, host='0.0.0.0',port=5001)
    #ホストをこの番号にするとローカルネットワーク内の他のデバイスからアクセス可能になる


