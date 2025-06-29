import pyperclip
from icalendar import Calendar, Event, vText
import datetime
import re
import time
import os
from dateutil import parser # 日付解析を強化するために追加

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

def create_and_open_ical(title, start_dt, description, location=""):
    """
    イベント情報からiCalファイルを生成し、デフォルトのカレンダーアプリで開きます。
    """
    cal = Calendar()
    cal.add('prodid', '-//My Python Calendar Helper//example.com//')
    cal.add('version', '2.0')

    event = Event()
    event.add('summary', vText(title))
    event.add('dtstart', start_dt)
    event.add('dtend', start_dt + datetime.timedelta(hours=1)) # デフォルトで1時間後
    event.add('description', vText(description))
    if location:
        event.add('location', vText(location))
    event.add('priority', 5)

    cal.add_component(event)

    # iCalファイルを一時的に保存
    file_name = f"temp_event_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.ics"
    with open(file_name, 'wb') as f:
        f.write(cal.to_ical())

    print(f"'{title}' のイベント情報を生成しました。カレンダーアプリで開きます...")

    # OSに応じた方法でファイルを開く
    try:
        if os.name == 'nt':  # Windows
            os.startfile(file_name)
        elif os.uname().sysname == 'Darwin':  # macOS
            os.system(f'open "{file_name}"')
        else:  # Linux (xdg-openが一般的)
            os.system(f'xdg-open "{file_name}"')
        print("カレンダーアプリケーションが起動しました。")
    except Exception as e:
        print(f"カレンダーアプリケーションの起動に失敗しました: {e}")
        print(f"手動で '{file_name}' を開いてカレンダーにインポートしてください。")
    
    # 必要であれば、ファイルを開いた後に削除することも可能ですが、
    # ユーザーがインポートを確認する時間を与えるため、ここでは残しておきます。
    # os.remove(file_name)

def monitor_clipboard_and_create_event(interval=1):
    """
    クリップボードの内容を監視し、変更があればイベントとして処理します。
    Ctrl+Cで終了できます。
    """
    last_clipboard_content = None
    print("クリップボードの監視を開始しました。テキストをコピーしてください。Ctrl+Cで終了します。")
    try:
        while True:
            current_clipboard_content = pyperclip.paste()
            if current_clipboard_content and current_clipboard_content != last_clipboard_content:
                print(f"\n--- クリップボードの内容が変更されました ({datetime.datetime.now().strftime('%H:%M:%S')}) ---")
                print(f"内容:\n{current_clipboard_content[:200]}...") # 長い場合は一部表示

                # イベント情報を解析
                title, start_dt, description, location = parse_event_from_text(current_clipboard_content)
                
                # 確認のため、解析結果を表示
                print(f"\n[解析結果]")
                print(f"  タイトル: {title}")
                print(f"  開始日時: {start_dt.strftime('%Y-%m-%d %H:%M')}")
                print(f"  説明: {description[:100]}...")
                print(f"  場所: {location if location else 'なし'}")

                # iCalファイルを生成して開く
                create_and_open_ical(title, start_dt, description, location)
                
                last_clipboard_content = current_clipboard_content
            time.sleep(interval)
    except KeyboardInterrupt:
        print("\nクリップボードの監視を終了します。")
    except pyperclip.PyperclipException as e:
        print(f"クリップボードの操作中にエラーが発生しました: {e}")
    except Exception as e:
        print(f"予期せぬエラーが発生しました: {e}")

if __name__ == "__main__":
    monitor_clipboard_and_create_event(interval=2) # 2秒ごとにチェック
    