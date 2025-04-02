import json
from datetime import datetime, timedelta
from icalendar import Calendar, Event
import re
import sys

def parse_time(time_str):
    """
    #YYYYMMDDThhmmss-YYYYMMDDThhmmss の形式を解析し、開始・終了のdatetimeオブジェクトを返す
    """
    today = datetime.today()  # 現在の日付を基準にする
    
    # とりあえず正規表現
    match = re.match(r'#(?:(\d{4}))?(?:(\d{2})(\d{2}))?T?(\d{2})?(\d{2})?(\d{2})?-?(?:(\d{4}))?(?:(\d{2})(\d{2}))?T?(\d{2})?(\d{2})?(\d{2})?', time_str)
    if not match:
        return None, None
    data = [list(match.groups()[0:6]), list(match.groups()[6:12])]
    
    # 調整
    for i in data:
        flag = False
        # 月
        if i[1] == None and i[0] != None:
            i[1] = i[0][0:2]
            flag = True
        # 日
        if i[2] == None and i[0] != None:
            i[2] = i[0][2:4]
            flag = True
        if flag:
            i[0] = None
    sy, sm, sd, sh, smin, ss = data[0]
    ey, em, ed, eh, emin, es = data[1]

    # Noneなら今日の日付追加
    year = int(sy) if sy else today.year
    month = int(sm) if sm else today.month
    day = int(sd) if sd else today.day
    
    # 作成
    start = datetime(year, month, day, int(sh or 0), int(smin or 0), int(ss or 0))
    if ey:
        end = datetime(int(ey), int(em), int(ed), int(eh or 0), int(emin or 0), int(es or 0))
    elif em and ed:
        end = datetime(today.year, int(em), int(ed), int(eh or 0), int(emin or 0), int(es or 0))
    elif eh:
        end = start.replace(hour=int(eh), minute=int(emin or 0), second=int(es or 0))
    else:
        # endに何もなければ+1したものをendにする
        if sm and sd and not sh:
            end = start + timedelta(days=1)
        elif sh:
            end = start + timedelta(hours=1)
        else:
            end = start
    
    return start, end

def time_delta(data):
    start, end = parse_time("#" + "-".join( map(lambda x: f"T{x}", data.split("-")) ))
    return str(int((end - start).total_seconds() / 60))

def format_string(template, data):
    """ {key} の形式のプレースホルダーをデータに基づいて置換 """
    def replace_match(match):
        key = match.group(1)
        if key.startswith("TID[") and key.endswith("]"):
            time_key = key[4:-1]
            if time_key in data:
                return time_delta(data[time_key])
        elif key.startswith("FOR[") and key.endswith("]"):
            content_key, fmt = key[4:].split("][")
            fmt = fmt.rstrip("]")
            res = ""
            # content_keyに#がついてたら#ignoreを実行する
            skip_flag = False
            if content_key[-1] == "#":
                skip_flag = True
                content_key = content_key[0:-1]
            # FORで回すkeyが存在するなら実行
            if content_key in data:
                # d: keyの中のデータ
                for d in data[content_key]:
                    # もし#ignoreがあって, 点数が-1ならスキップ
                    if "#ignore" in d[-1] and skip_flag:
                        continue
                    temp = fmt
                    # <i>を探す
                    for_match = re.findall(r'<(\d)>', temp)
                    # <i>を上書き, なければ空白
                    for i in set(for_match):
                        temp = temp.replace(f"<{i}>", str(d[int(i)]) if len(d) > int(i) else "")
                    # TIDがあるなら書き換える
                    tid_match = re.findall(r'TID\[(\d{4}-\d{4})\]', temp)
                    for i in set(tid_match):
                        temp = temp.replace(f"TID[{i}]", time_delta(i))
                    temp = temp.replace('(TID[]m)', "(外部利用)")
                    res += temp + "\n"
                return res
        return str(data.get(key, ""))
    
    return re.sub(r'\{([^}]+)\}', replace_match, template)

def generate_event(value, title, url_text, memo, desc):
    """ カレンダーにイベントを追加 """
    if isinstance(value, str) and value.startswith("#"):
        start, end = parse_time(value)
        if start and end:
            event = Event()
            event.add("summary", title)
            if url_text:
                # event.add("url", url_text)
                desc += "\n" + url_text
            if memo:
                desc += "\n" + memo
            event.add("description", desc)
            event.add("dtstart", start)
            event.add("dtend", end)
            return event
        print(f"Error: Time cant parse ({title})")
    # いずれも該当しないならNone
    return None

def generate_ics(format_json, data_json):
    # 大学ごとに作成
    for university in data_json:
        """ 書式JSONとデータJSONをもとにICSファイルを生成 """
        cal = Calendar()
        data_map = university.copy()
        univ_name = data_map["名称"]

        # カレンダーにメタデータ追加
        cal.add("X-WR-CALNAME", f"2025 {univ_name} カレンダー")
        cal.add("X-WR-CALDESC", f"{univ_name}カレンダー")
        cal.add("X-WR-RELCALID", data_map["url"])
        cal.add("PRODID", f"-//AutoGenerator//NONSGML v1.0//EN")
        
        # 入試形式ごと
        for exam_type, details in university.get("入試形式", {}).items():
            # 入試形式内のデータをdata_mapへ追記
            data_map.update(details)
            # key 入試形式 の調整
            data_map["入試形式"] = exam_type

            # フォーマットごと
            for key, format_entry in format_json.items():
                # keyが存在しないなら, スキップ
                if not key in details:
                    continue
                # イベント名を作成
                title = format_string(format_entry["title"], data_map)
                desc = format_string(format_entry.get("desc", ""), data_map)

                # 関連リンクの処理
                urls = []
                for link_key in format_entry.get("link", []):
                    # link_keyがpdfの中に存在する場合, 追加
                    if link_key in data_map["pdf"]["page"][exam_type]:
                        urls.append(link_key + ": " + data_map["pdf"]["link"] + "#page=" + str(data_map["pdf"]["page"][exam_type][link_key]))
                url_text = "\n".join(urls)

                # メモがあれば追加
                memo = ""
                try:
                    memo += "\n" + data_map["メモ"][exam_type][key]
                except KeyError:
                    pass
                
                # 日付を処理
                date_val = details[key]
                if isinstance(date_val, str):
                    temp = generate_event(date_val, title, url_text, memo, desc)
                    if temp:
                        cal.add_component(temp)
                elif isinstance(date_val, list):
                    # 複数ある場合, forで回す
                    for idx, v in enumerate(date_val):
                        title_idx = f"{title} ({idx+1})" if len(date_val) > 1 else title
                        temp = generate_event(v, title_idx, url_text, memo, desc)
                        if temp:
                            cal.add_component(temp)
            # END フォーマット
        # END 入試形式
        # ICSファイルとして出力
        with open(f"{univ_name}.ics", "wb") as f:
            f.write(cal.to_ical())

# サンプル実行
if __name__ == "__main__":
    format_filename = "format.json"
    univ_filename = "university_data.json"
    if len(sys.argv) >= 3:
        format_filename = sys.argv[2]
    if len(sys.argv) >= 2:
        univ_filename = sys.argv[1]
    # if len(sys.argv) == 1:
    #     print(f"Usage: python {sys.argv[0]} university_data.json (format.json)")
    #     exit()

    with open(format_filename, "r", encoding="utf-8") as f:
        format_data = json.load(f)
    
    with open(univ_filename, "r", encoding="utf-8") as f:
        university_data = json.load(f)
    
    generate_ics(format_data, university_data)
