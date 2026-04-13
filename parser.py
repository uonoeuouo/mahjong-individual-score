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

    # 1. テキスト解析
    for line in lines:
        # 正規表現: "名前 スコア [チョンボ]" を抽出
        match = re.match(r'^(.+?)[\s　]+([-\d\.]+)(?:[\s　]+(チョンボ))?$', line.strip())
        if match:
            raw_name = match.group(1).strip()
            normalized_name = jaconv.kata2hira(raw_name) 
            name = name_mapping.get(normalized_name, normalized_name) # 名前の表記揺れを変換
            try:
                score = float(match.group(2))
                game_data.append((name, score))
                if match.group(3) == 'チョンボ':
                    chombo_names.append(name)
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
    chombo_name_set = set(chombo_names)
    
    formatted_rows = []
    for rank, (name, score) in enumerate(ranked_data, 1):
        # [日付, 名前, スコア, 着順, チョンボ]
        chombo_mark = 'チョンボ' if name in chombo_name_set else ''
        formatted_rows.append([timestamp_str, name, score, rank, chombo_mark])

    return formatted_rows, None, chombo_names