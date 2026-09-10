"""
Discord Attendance Bot
-----------------------
Tracks voice-channel attendance for scheduled meetings and outputs a CSV:

    discord_username,attended,date,time_range
    Alice,true,2026-09-10,18:00-19:00
    Bob,true,2026-09-10,18:00-19:00
    Charlie,false,2026-09-10,18:00-19:00

Logic:
  1. /schedule_meeting creates a meeting tied to a voice channel + date + start/end time.
  2. At start time: anyone already in the VC -> attended = true.
  3. From start -> end: anyone who joins/moves into the VC -> attended = true
     (leaving does nothing, duration doesn't matter).
  4. At end time: anyone in the tracked list who never entered -> attended = false.
     The CSV is generated and appended to attendance.csv (accumulated across meetings).

Tracked members = everyone with a given role (default: everyone in the guild who
isn't a bot), passed in when scheduling the meeting.

Requirements: discord.py>=2.3, Python 3.10+
    pip install -U discord.py python-dotenv

Run:
    export DISCORD_TOKEN=your_token_here   (or put it in a .env file)
    python bot.py
"""

import asyncio
import csv
import os
from dataclasses import dataclass, field
from datetime import datetime, time as dtime
from pathlib import Path
from zoneinfo import ZoneInfo

import discord
from discord import app_commands
from discord.ext import commands

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

TOKEN = os.environ.get("DISCORD_TOKEN")
TIMEZONE = os.environ.get("BOT_TIMEZONE", "UTC")  # e.g. "America/Vancouver"
CSV_PATH = Path(os.environ.get("ATTENDANCE_CSV", "attendance.csv"))

TZ = ZoneInfo(TIMEZONE)

intents = discord.Intents.default()
intents.members = True
intents.voice_states = True

bot = commands.Bot(command_prefix="!", intents=intents)


# ---------------------------------------------------------------------------
# Meeting state
# ---------------------------------------------------------------------------

@dataclass
class Meeting:
    guild_id: int
    channel_id: int
    date_str: str          # "2026-09-10"
    time_range_str: str    # "18:00-19:00"
    start_dt: datetime
    end_dt: datetime
    tracked_member_ids: set[int]
    attended_ids: set[int] = field(default_factory=set)
    started: bool = False
    finished: bool = False


# channel_id -> Meeting (only meetings currently between "scheduled" and "finished")
active_meetings: dict[int, Meeting] = {}


def parse_meeting_datetimes(date_str: str, start_str: str, end_str: str) -> tuple[datetime, datetime]:
    date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
    start_t = datetime.strptime(start_str, "%H:%M").time()
    end_t = datetime.strptime(end_str, "%H:%M").time()

    start_dt = datetime.combine(date_obj, start_t, tzinfo=TZ)
    end_dt = datetime.combine(date_obj, end_t, tzinfo=TZ)

    if end_dt <= start_dt:
        raise ValueError("End time must be after start time.")

    return start_dt, end_dt


def get_tracked_members(guild: discord.Guild, role: discord.Role | None) -> set[int]:
    if role is not None:
        return {m.id for m in role.members if not m.bot}
    return {m.id for m in guild.members if not m.bot}


def write_csv_rows(meeting: Meeting, guild: discord.Guild) -> Path:
    """Append this meeting's attendance rows to the accumulated CSV."""
    file_exists = CSV_PATH.exists()

    rows = []
    for member_id in sorted(meeting.tracked_member_ids):
        member = guild.get_member(member_id)
        username = member.name if member else f"unknown_user_{member_id}"
        attended = member_id in meeting.attended_ids
        rows.append(
            {
                "discord_username": username,
                "attended": "true" if attended else "false",
                "date": meeting.date_str,
                "time_range": meeting.time_range_str,
            }
        )

    with CSV_PATH.open("a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f, fieldnames=["discord_username", "attended", "date", "time_range"]
        )
        if not file_exists:
            writer.writeheader()
        writer.writerows(rows)

    return CSV_PATH


# ---------------------------------------------------------------------------
# Meeting lifecycle
# ---------------------------------------------------------------------------

async def run_meeting(meeting: Meeting):
    """Waits for start, snapshots the VC, waits for end, finalizes + writes CSV."""
    guild = bot.get_guild(meeting.guild_id)
    if guild is None:
        return

    now = datetime.now(TZ)

    # --- Wait for / handle start ---
    delay = (meeting.start_dt - now).total_seconds()
    if delay > 0:
        active_meetings[meeting.channel_id] = meeting  # register before start so joins are caught early too
        await asyncio.sleep(delay)
    else:
        active_meetings[meeting.channel_id] = meeting

    channel = guild.get_channel(meeting.channel_id)
    if isinstance(channel, discord.VoiceChannel):
        for member in channel.members:
            if member.id in meeting.tracked_member_ids:
                meeting.attended_ids.add(member.id)
    meeting.started = True

    # --- Wait for end ---
    now = datetime.now(TZ)
    delay = (meeting.end_dt - now).total_seconds()
    if delay > 0:
        await asyncio.sleep(delay)

    # --- Finalize ---
    if meeting.finished:
        # Was cancelled while we were sleeping; skip CSV write.
        active_meetings.pop(meeting.channel_id, None)
        return
    meeting.finished = True
    active_meetings.pop(meeting.channel_id, None)
    csv_path = write_csv_rows(meeting, guild)

    # Notify in the channel that spawned this (best effort: post to system channel)
    summary_channel = guild.system_channel
    present = len(meeting.attended_ids)
    total = len(meeting.tracked_member_ids)
    text = (
        f"📋 Meeting on {meeting.date_str} ({meeting.time_range_str}) finished.\n"
        f"Attendance: {present}/{total} tracked members attended.\n"
        f"CSV updated: `{csv_path}`"
    )
    if summary_channel:
        try:
            await summary_channel.send(text)
        except discord.Forbidden:
            pass


# ---------------------------------------------------------------------------
# Voice state tracking
# ---------------------------------------------------------------------------

@bot.event
async def on_voice_state_update(member: discord.Member, before: discord.VoiceState, after: discord.VoiceState):
    if member.bot:
        return

    joined_channel_id = after.channel.id if after.channel else None
    if joined_channel_id is None:
        return  # only care about joins/moves INTO a channel

    meeting = active_meetings.get(joined_channel_id)
    if meeting is None or meeting.finished:
        return

    if member.id in meeting.tracked_member_ids:
        meeting.attended_ids.add(member.id)


# ---------------------------------------------------------------------------
# Slash commands
# ---------------------------------------------------------------------------

@bot.event
async def on_ready():
    await bot.tree.sync()
    print(f"Logged in as {bot.user} (tz={TIMEZONE}, csv={CSV_PATH.resolve()})")


@bot.tree.command(description="Schedule an attendance-tracked meeting on a voice channel.")
@app_commands.describe(
    channel="The voice channel to track",
    date="Meeting date, format YYYY-MM-DD",
    start="Start time, format HH:MM (24h)",
    end="End time, format HH:MM (24h)",
    role="Only track members with this role (defaults to all non-bot guild members)",
)
async def schedule_meeting(
    interaction: discord.Interaction,
    channel: discord.VoiceChannel,
    date: str,
    start: str,
    end: str,
    role: discord.Role | None = None,
):
    try:
        start_dt, end_dt = parse_meeting_datetimes(date, start, end)
    except ValueError as e:
        await interaction.response.send_message(f"⚠️ {e}", ephemeral=True)
        return

    if end_dt <= datetime.now(TZ):
        await interaction.response.send_message(
            "⚠️ That meeting's end time is already in the past.", ephemeral=True
        )
        return

    if channel.id in active_meetings:
        await interaction.response.send_message(
            "⚠️ There's already an active/scheduled meeting on that voice channel.",
            ephemeral=True,
        )
        return

    tracked = get_tracked_members(interaction.guild, role)
    time_range_str = f"{start}-{end}"

    meeting = Meeting(
        guild_id=interaction.guild_id,
        channel_id=channel.id,
        date_str=date,
        time_range_str=time_range_str,
        start_dt=start_dt,
        end_dt=end_dt,
        tracked_member_ids=tracked,
    )

    bot.loop.create_task(run_meeting(meeting))

    who = f"role **{role.name}**" if role else "all non-bot members"
    await interaction.response.send_message(
        f"✅ Meeting scheduled on **{channel.name}** for **{date} {time_range_str}** "
        f"(tz: {TIMEZONE}), tracking {who} ({len(tracked)} people)."
    )


@bot.tree.command(description="Check live attendance for a currently active meeting.")
@app_commands.describe(channel="The voice channel of the active meeting")
async def attendance_status(interaction: discord.Interaction, channel: discord.VoiceChannel):
    meeting = active_meetings.get(channel.id)
    if meeting is None:
        await interaction.response.send_message("No active meeting on that channel.", ephemeral=True)
        return

    guild = interaction.guild
    attended_names = sorted(
        (guild.get_member(mid).name if guild.get_member(mid) else str(mid))
        for mid in meeting.attended_ids
    )
    total = len(meeting.tracked_member_ids)
    lines = [
        f"Meeting: {meeting.date_str} {meeting.time_range_str}",
        f"Started: {meeting.started} | Finished: {meeting.finished}",
        f"Attended so far: {len(attended_names)}/{total}",
    ]
    if attended_names:
        lines.append("• " + "\n• ".join(attended_names))
    await interaction.response.send_message("\n".join(lines), ephemeral=True)


@bot.tree.command(description="Cancel a scheduled/active meeting without writing a CSV.")
@app_commands.describe(channel="The voice channel of the meeting to cancel")
async def cancel_meeting(interaction: discord.Interaction, channel: discord.VoiceChannel):
    meeting = active_meetings.pop(channel.id, None)
    if meeting is None:
        await interaction.response.send_message("No active meeting on that channel.", ephemeral=True)
        return
    meeting.finished = True  # prevents run_meeting from writing on wakeup
    await interaction.response.send_message("🛑 Meeting cancelled. No CSV was written.")


if __name__ == "__main__":
    if not TOKEN:
        raise SystemExit("Set the DISCORD_TOKEN environment variable (or .env file) before running.")
    bot.run(TOKEN)
