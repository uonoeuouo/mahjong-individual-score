import re

def parse_and_validate_message(message_content, timestamp_str, id_mapping=None):
    """
    メッセージテキストを解析し、整形されたデータとエラーログを返す
    戻り値: (game_rows, error_message)
    """
    if id_mapping is None:
        id_mapping = {}

    lines = message_content.strip().split('\n')
    game_data = [] # [(名前, スコア), ...]
    missing_members = [] # 辞書未登録のIDリスト

    # 1. テキスト解析
    for line in lines:
        # 正規表現: "名前 スコア" を抽出
        match = re.match(r'^(.+?)[\s　]+([-\d\.]+)$', line.strip())
        if match:
            user_id_str = match.group(1) # 文字列として取得
            score_str = match.group(2)
            
            # --- ID照合 ---
            if user_id_str in id_mapping:
                name = id_mapping[user_id_str]
            else:
                # 辞書にない場合
                name = f"未登録ID({user_id_str})"
                missing_members.append(user_id_str)

            try:
                score = float(match.group(2))
                game_data.append((name, score))
            except ValueError:
                continue

    # データがない、または人数不足の場合
    if len(game_data) != 4:
        return [], None # 4人揃ってない場合は無視、または別途エラー処理
    
    # 1. 未登録IDチェック (これ重要)
    if missing_members:
        return [], f"⚠️ 以下のメンバーがMembersシートに登録されていません: {', '.join(missing_members)}"

    # 2. 合計0チェック
    total_score = sum([s for _, s in game_data])
    if abs(total_score) > 0.1:
        return [], f"⚠️ 合計が0ではありません (合計: {total_score:.1f})"

    # 3. 順位計算とデータ整形
    # スコア降順でソート
    ranked_data = sorted(game_data, key=lambda x: x[1], reverse=True)
    
    formatted_rows = []
    for rank, (name, score) in enumerate(ranked_data, 1):
        # [日付, 名前, スコア, 着順]
        formatted_rows.append([timestamp_str, name, score, rank])

    return formatted_rows, None