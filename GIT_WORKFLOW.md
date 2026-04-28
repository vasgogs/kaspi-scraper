# GitHub and Restore Workflow

## What Goes To GitHub

Commit source code, scripts, service templates, documentation and safe examples.

Do not commit:

- `.env` with real tokens, phones, passwords or API keys
- `*.session` Telegram/Telethon sessions
- `logs/`, `RESULTS/`, `state/`, `user_data/`
- virtual environments: `venv/`, `.venv/`, `.venv_*`
- generated Excel/CSV reports and temporary files

## Daily Commands

Check what changed:

```bash
cd /home/vas/kaspi-scraper
git status
```

Save a version:

```bash
git add .
git commit -m "Describe what changed"
git push origin main
```

Get the latest version from GitHub:

```bash
git pull origin main
```

## Restore A File

See file history:

```bash
git log --oneline -- path/to/file.py
```

Restore one file from the last commit:

```bash
git restore path/to/file.py
```

Restore one file from an older commit:

```bash
git restore --source COMMIT_HASH -- path/to/file.py
```

Then commit the restoration:

```bash
git add path/to/file.py
git commit -m "Restore path/to/file.py"
```

## Restore The Whole Project To An Older Version

Safe way: create a new branch first, inspect it, then decide.

```bash
git switch -c restore-test COMMIT_HASH
```

If that version is correct, you can either keep working on this branch or copy the needed files back to `main`.

Hard reset is dangerous because it discards local changes. Use it only when you are sure:

```bash
git switch main
git reset --hard COMMIT_HASH
git push --force-with-lease origin main
```

## Restore On A New Server

```bash
git clone https://github.com/vasgogs/kaspi-scraper.git
cd kaspi-scraper
python3 -m venv venv
source venv/bin/activate
pip install -r telegram_webapp/requirements.txt
cp .env.example .env
```

Fill `.env` with real secrets and copy local data that is not stored in GitHub:

- mission/product CSV files
- Telethon `.session` file if phone broadcast should work without re-login
- any required historical `RESULTS/` or `state/` files

Restart services after restoring:

```bash
sudo systemctl restart kaspi-bot.service kaspi-mission.service kaspi-scraper-watchdog.service
```

## Useful Health Checks

```bash
systemctl status kaspi-bot.service kaspi-mission.service --no-pager
journalctl -u kaspi-mission.service -n 80 --no-pager
ls -lt RESULTS/mission_april_*.xlsx | head
```
