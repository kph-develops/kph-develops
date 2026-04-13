# HP Printer Usage Collector

A production-ready Python application that connects to HP printers via their
embedded web interface, scrapes page-count and toner-level data, and emails a
formatted monthly report.

---

## Features

- Scrapes **Total Page Count** and **CMYK toner levels** from HP embedded web pages
- Supports **multiple printers** in a single run — all included in one report
- **Low-toner alerts** (configurable threshold, default 20 %)
- **HTML + plain-text** multipart email via SMTP with STARTTLS
- Appends every run to a **CSV history file** for trend analysis
- **Rotating log file** with console mirroring
- Zero-friction scheduling via **Windows Task Scheduler** or **cron** (Linux/WSL)

---

## Project Structure

```
hp_printer_collector/
├── hp_printer_collector/        # Python package
│   ├── __init__.py
│   ├── logger_setup.py          # Rotating file + console logging
│   ├── scraper.py               # HTTP fetch and HTML parsing
│   ├── storage.py               # CSV history writer
│   └── email_reporter.py        # Report builder and SMTP sender
├── main.py                      # Entry point / CLI
├── config.example.yaml          # Annotated sample configuration
├── config.yaml                  # Your live config (gitignored)
├── requirements.txt
└── README.md
```

---

## Quick Start

### 1. Prerequisites

- **Python 3.9 or later** — download from <https://www.python.org/downloads/>
  - On Windows tick **"Add Python to PATH"** during installation
- Network access to the HP printer's web interface (default port 80)

### 2. Clone / download the project

```bash
git clone <repo-url> hp_printer_collector
cd hp_printer_collector
```

### 3. Create a virtual environment (recommended)

**Windows (Command Prompt)**
```cmd
python -m venv venv
venv\Scripts\activate
```

**Windows (PowerShell)**
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

**Linux / macOS / WSL**
```bash
python3 -m venv venv
source venv/bin/activate
```

### 4. Install dependencies

```bash
pip install -r requirements.txt
```

### 5. Configure

```bash
cp config.example.yaml config.yaml
```

Open `config.yaml` in any text editor and fill in:

| Section | Key | Description |
|---------|-----|-------------|
| `printers` | `ip` | IP address of the HP printer |
| `printers` | `name` | Human-readable label (optional) |
| `smtp` | `host` | SMTP server hostname |
| `smtp` | `port` | Usually `587` (STARTTLS) or `25` (local relay) |
| `smtp` | `username` | SMTP login (leave blank for unauthenticated relays) |
| `smtp` | `password` | SMTP password or Gmail App Password |
| `email` | `recipients` | List of To: addresses |
| `alerts` | `toner_low_threshold` | Alert threshold in % (default `20`) |

> **Gmail users:** You must create an **App Password** at
> <https://myaccount.google.com/apppasswords> and use that instead of your
> regular account password.  Two-Factor Authentication must be enabled.

### 6. Test the connection

```bash
python main.py --no-email
```

This collects data and writes to the CSV without sending an email — useful for
verifying printer connectivity and config before scheduling.

### 7. Full run (collect + email)

```bash
python main.py
```

---

## Command-Line Options

```
python main.py [OPTIONS]

  --config PATH    Path to the YAML config file
                   (default: config.yaml next to main.py)
  --no-email       Collect data and save CSV, but skip sending email
  --no-csv         Do not append results to the CSV history file
```

**Exit codes**

| Code | Meaning |
|------|---------|
| `0` | All printers collected, email sent |
| `1` | One or more printers failed, or email failed |
| `2` | Configuration error |

---

## Scheduling

### Windows Task Scheduler (recommended for Windows)

1. Open **Task Scheduler** → *Create Basic Task*
2. **Name:** `HP Printer Monthly Report`
3. **Trigger:** Monthly → Day `1` → time `07:00`
4. **Action:** *Start a program*
   - Program: `C:\path\to\hp_printer_collector\venv\Scripts\python.exe`
   - Arguments: `C:\path\to\hp_printer_collector\main.py`
   - Start in: `C:\path\to\hp_printer_collector`
5. **Finish** → tick *Open Properties dialog* → *Settings* tab →
   tick *Run task as soon as possible after a scheduled start is missed*
6. Optionally tick *Run whether user is logged on or not* for headless operation

**PowerShell one-liner to register the task:**

```powershell
$action  = New-ScheduledTaskAction `
    -Execute "C:\path\to\venv\Scripts\python.exe" `
    -Argument "C:\path\to\hp_printer_collector\main.py" `
    -WorkingDirectory "C:\path\to\hp_printer_collector"

$trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At "07:00"

Register-ScheduledTask `
    -TaskName "HP Printer Monthly Report" `
    -Action $action `
    -Trigger $trigger `
    -RunLevel Highest `
    -Description "Collect HP printer stats and email monthly report"
```

---

### Cron (Linux / macOS / WSL)

Add to your crontab (`crontab -e`):

```cron
# Run at 07:00 on the 1st of every month
0 7 1 * * /path/to/hp_printer_collector/venv/bin/python /path/to/hp_printer_collector/main.py >> /path/to/hp_printer_collector/cron.log 2>&1
```

---

## Output Files

| File | Description |
|------|-------------|
| `printer_collector.log` | Rotating log (5 MB × 5 files). Path configured in `config.yaml`. |
| `printer_history.csv` | One row per printer per run. Columns: `timestamp`, `printer_name`, `printer_ip`, `page_count`, `toner_black`, `toner_cyan`, `toner_yellow`, `toner_magenta`, `error`. |

---

## Parsed HTML Elements

| Data Point | Page | Element |
|-----------|------|---------|
| Total page count | `UsagePage` | `<td id="UsagePage.EquivalentImpressionsTable.Total.Total">` |
| Black toner | `SuppliesStatus` | `id="BlackCartridge1-Header_Level"` |
| Cyan toner | `SuppliesStatus` | `id="CyanCartridge1-Header_Level"` |
| Yellow toner | `SuppliesStatus` | `id="YellowCartridge1-Header_Level"` |
| Magenta toner | `SuppliesStatus` | `id="MagentaCartridge1-Header_Level"` *(falls back to regex scan of `id="MagentaCartridge1-Header"` section)* |

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| `Cannot connect to printer at 192.168.x.x` | Wrong IP or firewall | Verify IP via printer display; check network routing |
| `Page count element not found` | Different firmware version | Inspect the page source and update the element ID in `scraper.py` |
| `SMTP authentication failed` | Wrong credentials | For Gmail use an App Password; check 2FA is enabled |
| Email received but toner shows `N/A` | Supplies page unavailable | Try opening `http://<IP>/hp/device/InternalPages/Index?id=SuppliesStatus` in a browser |
| Windows Task Scheduler task never ran | Task not running as correct user | In task properties set *Run whether user is logged on or not* and supply credentials |

---

## Security Notes

- **Store credentials safely.** `config.yaml` contains your SMTP password.
  Set filesystem permissions so only the service account can read it:
  ```cmd
  icacls config.yaml /inheritance:r /grant:r "%USERNAME%:R"
  ```
- The application disables SSL certificate verification when connecting to
  printers because HP embedded web servers use self-signed certificates.
  This is acceptable for a trusted internal network.
- Do not commit `config.yaml` to source control — it is listed in `.gitignore`.

---

## License

MIT
