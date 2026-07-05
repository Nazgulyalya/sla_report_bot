# SLA Report Bot

Telegram bot that automates the generation and delivery of SLA (Service Level 
Agreement) reports for infrastructure services. Built to replace a manual 
reporting process during my work as a Monitoring Specialist.

## Problem
SLA reports were assembled by hand every day — pulling metrics, formatting, 
and sending them out. Slow and error-prone.

## What it does
- Queries availability & performance metrics from a PostgreSQL database
- Aggregates them into a clean SLA summary (uptime %, incidents, etc.)
- Sends the formatted report automatically to a Telegram channel on schedule

## Impact
Reduced manual report preparation time by ~80%.

## Tech stack
Python · PostgreSQL (SQL) · Telegram Bot API · [scheduler — cron / schedule / APScheduler]

## Run
```bash
pip install -r requirements.txt
# add your credentials to .env (BOT_TOKEN, DB_URL)
python [main.py]
```
