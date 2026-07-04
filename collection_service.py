from parser import parse_and_validate_message
from collection_result import CollectionResult
from datetime import timedelta, timezone
JST = timezone(timedelta(hours=9))

async def run_collection_process(channel, client, sheet_handler, completion_message):
    try:
        name_mapping = sheet_handler.get_name_mapping()

        last_checkpoint = await _find_last_checkpoint(channel, client.user)
        history = _build_history_iterator(channel, last_checkpoint)

        collection_result = await _collect_history_rows(history, name_mapping)

        result_msg = _write_results_and_build_message(
            sheet_handler,
            collection_result,
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
    collection_result = CollectionResult()

    async for msg in history:
        if msg.author.bot or msg.content.startswith('!'):
            continue

        timestamp = msg.created_at.astimezone(JST).strftime('%Y/%m/%d %H:%M')
        rows, error, chombo_names = parse_and_validate_message(msg.content, timestamp, name_mapping)

        if error:
            collection_result.add_error(timestamp, error)
            continue

        if not rows:
            continue

        collection_result.add_game(rows, msg.created_at, chombo_names)

    return collection_result


def _write_results_and_build_message(sheet_handler, collection_result, completion_message):
    if collection_result.has_rows():
        sheet_handler.append_game_data(collection_result.raw_rows)
        daily_sheet_payload = collection_result.daily_sheet_payload()
        if daily_sheet_payload:
            sheet_handler.record_daily_activities_batch(daily_sheet_payload)
        sheet_handler.record_stats_chombo_counts()
        result_msg = f"{completion_message}\n追加件数: {collection_result.game_count} 試合"
    else:
        result_msg = "✅ 集計は行われませんでした。"

    if collection_result.error_logs:
        result_msg += "\n\n【エラー報告】\n" + "\n".join(collection_result.error_logs)

    return result_msg
