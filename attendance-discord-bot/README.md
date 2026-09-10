# Discord Attendance Bot

Tracks who attends a voice-channel meeting and writes/append a CSV:

```csv
discord_username,attended,date,time_range
Alice,true,2026-09-10,18:00-19:00
Bob,true,2026-09-10,18:00-19:00
Charlie,false,2026-09-10,18:00-19:00
```

## How it works

1. `/schedule_meeting` registers a meeting on a voice channel with a date + start/end time.
2. At start time, anyone already in the VC is marked `attended = true`.
3. Between start and end, anyone who joins or moves into the VC is marked `attended = true`
   (duration doesn't matter, and leaving does **not** undo attendance).
4. At end time, anyone tracked who never entered is marked `attended = false`, and the
   attendance rows for that meeting are appended to `attendance.csv`.

Multiple meetings accumulate into the same CSV file (one meeting's rows per run), so you
get a running attendance log across meetings — matching the format you described.

## Setup

```bash
pip install -r requirements.txt
```

Create a `.env` file (or set environment variables directly):

```
DISCORD_TOKEN=your_bot_token_here
BOT_TIMEZONE=America/Vancouver   # optional, defaults to UTC
ATTENDANCE_CSV=attendance.csv    # optional, defaults to ./attendance.csv
```

### Bot permissions / intents

In the Discord Developer Portal, under your bot's **Bot** settings, enable:
- `SERVER MEMBERS INTENT`
- `PRESENCE INTENT` is not required, but `SERVER MEMBERS INTENT` and the ability to see
  voice states is required (voice states are visible via the standard Voice States intent,
  enabled in code — no extra portal toggle needed for that one).

Invite the bot with the `applications.commands` and `bot` scopes, and at least these
permissions: `View Channels`, `Connect` (not required to actually join voice — the bot
only *observes* voice state events, it never joins the call), `Send Messages`.

## Run

```bash
python bot.py
```

## Commands (slash commands)

- **/schedule_meeting** `channel` `date` `start` `end` `[role]`
  Schedule a meeting. `date` is `YYYY-MM-DD`, `start`/`end` are `HH:MM` 24-hour, in the
  bot's configured timezone. If `role` is omitted, every non-bot member of the guild is
  tracked; if given, only members with that role are tracked.

- **/attendance_status** `channel`
  Peek at live attendance for a meeting that's currently in progress.

- **/cancel_meeting** `channel`
  Cancel a scheduled or in-progress meeting without writing any CSV rows for it.

## Notes / limitations

- Meeting schedules live in memory only. If the bot restarts before a meeting's end time,
  that meeting's tracking is lost — reschedule it after restart if needed. For
  production use, persist `active_meetings` to disk/DB and reload on startup.
- Usernames in the CSV use the Discord `member.name` (account username, not the
  per-server nickname/display name). Swap in `member.display_name` in `write_csv_rows`
  if you'd rather log server nicknames.
- Only one meeting can be scheduled per voice channel at a time (by design, to avoid
  overlapping trackers on the same channel).
