from flask import Flask, request, send_file, render_template_string 
import pyperclip
from icalendar import Calendar,Event, vText
import datetime
import re
import os
from dateutil import parser
import io

app= Flask(__name__)

import datetime
from dateutil import parser
import re

def parse_event_from_text(text_content):
    title = "新しいカレンダーイベント"
    location = ""
    duration = None # 新しくdurationを追加

    now = datetime.datetime.now()
    # デフォルトの開始日時を、現在の日付の午前9時に設定
    # もし現在時刻が午前9時を過ぎていたら、翌日の午前9時にする
    default_start_dt = now.replace(hour=9, minute=0, second=0, microsecond=0)
    if now.hour >= 9:
        default_start_dt += datetime.timedelta(days=1)
    
    start_dt = default_start_dt # 初期値を設定

    # --- 1. 日付の解析 ---
    # まずはYYYY/MM/DD, MM/DD, YYYY年MM月DD日 などの日付パターンを抽出
    date_patterns = [
        r'(\d{4}[/\-年]\d{1,2}[/\-月]\d{1,2}日?)', # YYYY/MM/DD or YYYY-MM-DD
        r'(\d{1,2}[/\-月]\d{1,2}日?)',          # MM/DD or MM-DD
        r'(\d{1,2}月\d{1,2}日)',                # MM月DD日
    ]
    
    found_date_in_text = False
    for pattern in date_patterns:
        date_match = re.search(pattern, text_content)
        if date_match:
            date_str = date_match.group(1)
            try:
                # MM/DD 形式の場合は dayfirst=False を優先
                if re.match(r'\d{1,2}[/\-月]\d{1,2}日?', date_str):
                    parsed_dt = parser.parse(date_str, dayfirst=False, fuzzy=True)
                else:
                    parsed_dt = parser.parse(date_str, fuzzy=True) # その他の形式はデフォルトで
                
                # 解析された日付でstart_dtの日付部分を更新
                start_dt = start_dt.replace(year=parsed_dt.year, 
                                            month=parsed_dt.month, 
                                            day=parsed_dt.day)
                found_date_in_text = True
                break # 最初に見つかった日付を採用
            except parser.ParserError:
                continue # 解析失敗したら次のパターンを試す
    
    # 相対日付の処理 (もし上記で日付が見つからなければ)
    if not found_date_in_text:
        if "明日" in text_content:
            start_dt = now + datetime.timedelta(days=1)
            start_dt = start_dt.replace(hour=default_start_dt.hour, minute=default_start_dt.minute, second=0, microsecond=0)
            found_date_in_text = True
        elif "明後日" in text_content:
            start_dt = now + datetime.timedelta(days=2)
            start_dt = start_dt.replace(hour=default_start_dt.hour, minute=default_start_dt.minute, second=0, microsecond=0)
            found_date_in_text = True
        elif "今日" in text_content:
            start_dt = now.replace(hour=default_start_dt.hour, minute=default_start_dt.minute, second=0, microsecond=0)
            found_date_in_text = True
        # "昨日" や "一昨日" はカレンダー作成では通常使わないため、ここでは含めません。

    # --- 2. 時刻の解析 ---
    time_match = re.search(r'(\d{1,2}:\d{2}(?:[APap][Mm])?|\d{1,2}時\d{2}分?|\d{1,2}時)', text_content)
    if time_match:
        try:
            time_str = time_match.group(1).replace('時', ':').replace('分', '')
            if 'am' in time_str.lower() or 'pm' in time_str.lower():
                parsed_time = datetime.datetime.strptime(time_str, '%I:%M%p').time()
            elif ':' in time_str:
                parsed_time = datetime.datetime.strptime(time_str, '%H:%M').time()
            else: # 例: "10時" の場合
                parsed_time = datetime.datetime.strptime(time_str + ':00', '%H:%M').time()

            # 時刻が見つかった場合は、解析された時刻でstart_dtの時刻部分を更新
            start_dt = start_dt.replace(hour=parsed_time.hour, minute=parsed_time.minute, second=0, microsecond=0)

            # もし解析された日時が現在時刻より過去で、かつ日付が今日の場合、翌日にする
            # (これは元のロジックを維持。テキストに例えば「9時」とだけあり、現在が10時の場合、翌日の9時にする)
            if start_dt < datetime.datetime.now() and start_dt.date() == datetime.datetime.now().date():
                start_dt += datetime.timedelta(days=1)

        except ValueError:
            pass # 解析失敗時は設定済みのstart_dt (デフォルトまたは日付解析結果) を使用

    # --- 3. 期間の解析 (duration) ---
    end_dt = None
    # 例: "7/1~7/2 10:00~12:00" のような形式を解析
    # 日付範囲と時刻範囲の抽出
    # 複数のパターンを試す
    datetime_range_patterns = [
        r'(\d{1,2}[/\-月]\d{1,2}日?)~?(\d{1,2}[/\-月]\d{1,2}日?)?\s*(\d{1,2}:\d{2})~?(\d{1,2}:\d{2})?', # 7/1~7/2 10:00~12:00
        r'(\d{1,2}:\d{2})~?(\d{1,2}:\d{2})?', # 10:00~12:00 (同日内)
        r'(\d{1,2}[/\-月]\d{1,2}日?)~?(\d{1,2}[/\-月]\d{1,2}日?)', # 7/1~7/2 (終日イベント)
    ]

    for pattern in datetime_range_patterns:
        datetime_range_match = re.search(pattern, text_content)
        if datetime_range_match:
            try:
                # グループの取得と整形
                g = datetime_range_match.groups()
                
                # 開始日時を再確認 (既にstart_dtで解析済みだが、期間解析でより正確な情報が得られる場合があるため)
                # グループの数でどのパターンにマッチしたか判断
                if len(g) >= 4 and g[0] and g[2]: # 7/1~7/2 10:00~12:00 のようなパターン
                    start_date_str = g[0]
                    start_time_str = g[2]
                    parsed_start_dt_from_range = parser.parse(start_date_str + " " + start_time_str, dayfirst=False, fuzzy=True)
                    start_dt = start_dt.replace(year=parsed_start_dt_from_range.year,
                                                month=parsed_start_dt_from_range.month,
                                                day=parsed_start_dt_from_range.day,
                                                hour=parsed_start_dt_from_range.hour,
                                                minute=parsed_start_dt_from_range.minute,
                                                second=0, microsecond=0)
                elif len(g) >= 2 and g[0] and not g[1]: # 10:00~12:00 のような時刻のみのパターン
                    start_time_str = g[0]
                    parsed_start_time = datetime.datetime.strptime(start_time_str, '%H:%M').time()
                    start_dt = start_dt.replace(hour=parsed_start_time.hour, minute=parsed_start_time.minute, second=0, microsecond=0)
                    # 時刻が過去なら翌日にするロジックは、上記時刻解析部分で既に考慮済み

                # 終了日時の解析
                if len(g) >= 4 and g[1] and g[3]: # 7/1~7/2 10:00~12:00 のようなパターン
                    end_date_str = g[1]
                    end_time_str = g[3]
                    parsed_end_dt = parser.parse(end_date_str + " " + end_time_str, dayfirst=False, fuzzy=True)
                    # もし終了日が開始日より前なら、終了日を翌年に調整など
                    if parsed_end_dt < start_dt:
                        parsed_end_dt = parsed_end_dt.replace(year=start_dt.year + 1) # 簡単な調整
                    end_dt = parsed_end_dt
                elif len(g) >= 2 and g[1] and not g[0]: # 10:00~12:00 のような時刻のみのパターン
                    end_time_str = g[1]
                    parsed_end_time = datetime.datetime.strptime(end_time_str, '%H:%M').time()
                    end_dt = start_dt.replace(hour=parsed_end_time.hour, minute=parsed_end_time.minute, second=0, microsecond=0)
                    if end_dt < start_dt: # 例: 10:00~09:00 の場合
                        end_dt += datetime.timedelta(days=1)
                elif len(g) >= 2 and g[1] and not g[2]: # 7/1~7/2 のような日付のみのパターン (終日イベント)
                    end_date_str = g[1]
                    parsed_end_dt = parser.parse(end_date_str, dayfirst=False, fuzzy=True)
                    end_dt = parsed_end_dt.replace(hour=23, minute=59, second=59) # 終日の場合、終了日の終わりに設定
                    if end_dt < start_dt:
                        end_dt = end_dt.replace(year=start_dt.year + 1)

                if end_dt:
                    time_delta = end_dt - start_dt
                    total_seconds = int(time_delta.total_seconds())
                    
                    # durationは最低1分を設定
                    if total_seconds <= 0:
                        duration = "1m"
                    else:
                        days = total_seconds // (24 * 3600)
                        hours = (total_seconds % (24 * 3600)) // 3600
                        minutes = (total_seconds % 3600) // 60
                        seconds = total_seconds % 60
                        
                        duration_parts = []
                        if days > 0: duration_parts.append(f"{days}d")
                        if hours > 0: duration_parts.append(f"{hours}h")
                        if minutes > 0: duration_parts.append(f"{minutes}m")
                        if seconds > 0: duration_parts.append(f"{seconds}s")
                        
                        duration = "".join(duration_parts)
                
                break # 期間解析に成功したらループを抜ける

            except Exception as e:
                # print(f"Duration parsing error: {e}") # デバッグ用
                duration = None # 解析失敗時は期間なし
                continue # 次のパターンを試す

    # 4. タイトルの解析 (最も重要な情報)
    # 日時や期間情報を削除してからタイトルを抽出
    cleaned_text_for_title = text_content
    # 日時範囲の正規表現パターンを結合して一度に削除
    all_datetime_patterns_for_removal = r'|'.join([p.replace('(', '(?:') for p in date_patterns + datetime_range_patterns])
    cleaned_text_for_title = re.sub(all_datetime_patterns_for_removal, '', cleaned_text_for_title).strip()
    
    title_keywords = ["会議", "打ち合わせ", "ミーティング", "予約", "イベント", "リマインダー", "予定"] # 「予定」も追加
    found_title_from_keyword = False
    for keyword in title_keywords:
        if keyword in cleaned_text_for_title:
            # キーワードより前の部分をタイトルにする
            parts = cleaned_text_for_title.split(keyword, 1)
            if parts[0].strip():
                title = parts[0].strip() # キーワード自体はタイトルに含めない
            else: # キーワードが先頭にある場合など
                title = keyword # キーワード自体をタイトルにする
            found_title_from_keyword = True
            break
    
    if not found_title_from_keyword:
        # 日時情報などを取り除いた残りのテキストをタイトルにする
        cleaned_text = re.sub(r'(\d{4}[年/]\d{1,2}[/\-月]\d{1,2}日?|\d{1,2}[/\-月]\d{1,2}日?|\d{1,2}月\d{1,2}日|今日|明日|明後日|昨日|一昨日|\d{1,2}:\d{2}(?:[APap][Mm])?|\d{1,2}時\d{2}分?|\d{1,2}時|\d{1,2}[/\-月]\d{1,2}日?~?\d{1,2}[/\-月]\d{1,2}日?\s*\d{1,2}:\d{2}~?\d{1,2}:\d{2})', '', text_content)
        cleaned_text = cleaned_text.replace('\n', ' ').strip()
        if cleaned_text:
            title = cleaned_text[:100] + '...' if len(cleaned_text) > 100 else cleaned_text
        else:
            title = "クリップボードからのイベント"

    # 5. 説明 (残りのテキスト全て、または特定の情報)
    description = text_content # 全体を説明とする

    # 6. 場所の解析 (例: "場所：〇〇", "@〇〇" など)
    location_match = re.search(r'(場所|@|にて)[:：]?\s*(.+?)(?=\n|$)', text_content)
    if location_match:
        location = location_match.group(2).strip()

    return title, start_dt, duration, description, location

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


