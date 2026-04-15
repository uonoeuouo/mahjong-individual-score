import re
import jaconv

def parse_and_validate_message(message_content, timestamp_str, name_mapping=None):
    """
    メッセージテキストを解析し、整形されたデータとエラーログを返す
    戻り値: (game_rows, error_message, chombo_names)
    """
    lines = message_content.strip().split('\n')
    game_data = [] # [(名前, スコア), ...]
    chombo_names = []
    chombo_count_map = {}
    
    # 1. テキスト解析
    for line in lines:
        # 正規表現: "名前 スコア [チョンボ][回数]" を抽出
        match = re.match(r'^(.+?)[\s　]+([-\d\.]+)(?:[\s　]+(チョンボ)(?:[\s　]*(\d+))?)?$', line.strip())

        if match:
            raw_name = match.group(1).strip()
            normalized_name = jaconv.kata2hira(raw_name) 
            name = name_mapping.get(normalized_name, normalized_name) # 名前の表記揺れを変換

            try:
                score = float(match.group(2))
                game_data.append((name, score))

                chombo_keyword = match.group(3)
                chombo_num_str = match.group(4)

                if chombo_keyword:
                    if chombo_num_str:
                        # 全角数字を半角に正規化してから数値化
                        normalized_num = jaconv.z2h(chombo_num_str, digit=True, ascii=False, kana=False)
                        count = int(normalized_num)
                    else:
                        # 「チョンボ」のみで数字なしは1回
                        count = 1

                    chombo_count_map[name] = chombo_count_map.get(name, 0) + count
                    chombo_names.extend([name] * count)
            except ValueError:
                continue

    # データがない、または人数不足の場合
    if len(game_data) != 4:
        return [], None, [] # 4人揃ってない場合は無視、または別途エラー処理

    # 2. 合計0チェック
    total_score = sum([s for _, s in game_data])
    if abs(total_score) > 0.1:
        return [], f"⚠️ 合計が0ではありません (合計: {total_score:.1f})", []

    # 3. 順位計算とデータ整形
    # スコア降順でソート
    ranked_data = sorted(game_data, key=lambda x: x[1], reverse=True)
    
    formatted_rows = []
    for rank, (name, score) in enumerate(ranked_data, 1):
        # [日付, 名前, スコア, 着順, チョンボ]
        chombo_count = chombo_count_map.get(name, 0)
        if chombo_count <= 0:
            chombo_mark = ''
        elif chombo_count == 1:
            chombo_mark = 'チョンボ'
        else:
            chombo_mark = f'チョンボ{chombo_count}'
        formatted_rows.append([timestamp_str, name, score, rank, chombo_mark])

    return formatted_rows, None, chombo_names