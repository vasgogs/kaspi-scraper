## Kaspi Scraper Automation

Follow these steps once and the scraper will upload daily results to the SharePoint workbook automatically.

### 1. Install dependencies

```bash
cd "/Users/vas/Documents/stada - all docs/kaspi scraper"
pip install -r requirements.txt  # if you already track deps
pip install playwright office365-rest-python-client pandas openpyxl
playwright install
```

### 2. Configure SharePoint access

1. Duplicate `.env.example` to `.env`.
2. Fill in your SharePoint site URL, file relative URL, and credentials.
3. Keep `.env` private—`run_kaspi.sh` sources it automatically.

### 3. Test a manual run

```bash
cd "/Users/vas/Documents/stada - all docs/kaspi scraper"
chmod +x run_kaspi.sh
./run_kaspi.sh
```

Confirm:
- A new Excel file is created under `RESULTS/`.
- The SharePoint workbook shows the appended rows.

### 4. Schedule daily execution (cron example)

Run `crontab -e` and add:

```
0 8 * * * /Users/vas/Documents/stada\ -\ all\ docs/kaspi\ scraper/run_kaspi.sh >> /Users/vas/Documents/kaspi_cron.log 2>&1
```

This runs every day at 08:00, uses the `.env` credentials, and logs output to `~/Documents/kaspi_cron.log`.

### 5. Monitoring

- Check the log file periodically for failures.
- Replace credentials immediately if compromised and update `.env`.
