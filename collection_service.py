from parser import parse_and_validate_message


async def run_collection_process(channel, client, sheet_handler, completion_message):
    try:
        name_mapping = sheet_handler.get_name_mapping()

        last_checkpoint = await _find_last_checkpoint(channel, client.user)
        history = _build_history_iterator(channel, last_checkpoint)

        all_rows_to_add, error_logs, daily_batches = await _collect_history_rows(history, name_mapping)

        result_msg = _write_results_and_build_message(
            sheet_handler,
            all_rows_to_add,
            daily_batches,
            error_logs,
            completion_message,
        )
        await channel.send(result_msg)

    except Exception as e:
        await channel.send(f"❌ エラーが発生しました: {e}")
        print(f"Error: {e}")


async def _find_last_checkpoint(channel, bot_user):
    last_checkpoint = None
    async for msg in channel.history(limit=100):
        if msg.author == bot_user and msg.content.startswith("✅ 集計完了"):
            last_checkpoint = msg
            break
    return last_checkpoint


def _build_history_iterator(channel, last_checkpoint):
    if last_checkpoint:
        print(f"前回の完了地点: {last_checkpoint.created_at}")
        return channel.history(after=last_checkpoint, limit=None, oldest_first=True)

    print("全件読み込みモード")
    return channel.history(limit=200, oldest_first=True)


async def _collect_history_rows(history, name_mapping):
    all_rows_to_add = []
    error_logs = []
    daily_batches = {}
    async for msg in history:
        if msg.author.bot or msg.content.startswith('!'):
            continue

        timestamp = msg.created_at.strftime('%Y/%m/%d %H:%M')
        rows, error, chombo_names = parse_and_validate_message(msg.content, timestamp, name_mapping)

        if error:
            error_logs.append(f"⚠️ {timestamp} の投稿: {error}")
            continue

        if not rows:
            continue

        all_rows_to_add.extend(rows)

        sheet_date = msg.created_at.strftime('%Y%m%d')
        game_data = [(r[1], r[2], r[3], r[4]) for r in rows]
        batch = daily_batches.setdefault(sheet_date, {'games': [], 'chombo_counts': {}})
        batch['games'].append(game_data)

        for name in chombo_names:
            batch['chombo_counts'][name] = batch['chombo_counts'].get(name, 0) + 1

    return all_rows_to_add, error_logs, daily_batches


def _write_results_and_build_message(sheet_handler, all_rows_to_add, daily_batches, error_logs, completion_message):
    if all_rows_to_add:
        sheet_handler.append_game_data(all_rows_to_add)
        if daily_batches:
            sheet_handler.record_daily_activities_batch(daily_batches)
        sheet_handler.record_stats_chombo_counts()
        result_msg = f"{completion_message}\n追加件数: {len(all_rows_to_add)//4} 試合"
    else:
        result_msg = "✅ 集計は行われませんでした。"

    if error_logs:
        result_msg += "\n\n【エラー報告】\n" + "\n".join(error_logs)

    return result_msg
