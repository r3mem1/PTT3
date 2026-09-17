---
name: run
description: Launch or restart the presentation_bot Telegram bot (python bot.py) and confirm it started polling cleanly.
---

This project is a long-running Telegram bot (aiogram), not a request/response server —
"running" it means starting `python bot.py` and watching it long-poll Telegram, not
hitting a URL.

To (re)start it:

1. Check for an already-running instance first — starting a second one causes a
   Telegram `Conflict: terminated by other getUpdates request` error, since only one
   process can long-poll a given bot token at a time. If one is already running (from
   a previous background task in this session), stop it before starting a new one —
   this is required whenever the bot's Python source changed, since Python does not
   hot-reload a running process.
2. Start `python bot.py` from the repo root with `run_in_background: true`.
3. Wait for the log lines `Start polling` and `Run polling for bot @<username>` — that
   confirms Telegram accepted the long-poll and there's no token/network problem. No
   output within a few seconds, or a traceback instead, means startup failed — read
   the full output file, don't assume success from silence.
4. To verify an actual generation end-to-end, either wait for a real Telegram message
   (the user sends one) and tail the log for the `Update id=... is handled. Duration
   ...` line — a duration under ~1s for a topic message means it errored out before
   calling OpenRouter (check for a traceback just above it); 15-90+ seconds is normal
   for a real generation — or drive the pipeline directly without Telegram (see the
   `render-check` skill for the no-AI-call rendering path, and the `prompt-tester`
   agent for the with-AI-call path).
