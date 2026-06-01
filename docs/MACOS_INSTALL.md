# 🍎 MinervaDB XL View — macOS Installation Guide

Complete guide for installing and running MinervaDB XL View on macOS (Intel & Apple Silicon).

> **Note on Excel for Mac:** Microsoft Excel for Mac does **not** support VBA ADODB/ODBC connectivity the same way Windows does. The recommended approach on macOS is to use the **Python ETL pipeline** to extract data from PostgreSQL and write it into Excel files (.xlsx), then open those in Excel for Mac. Alternatively, you can run the full VBA experience inside a Windows VM (Parallels Desktop, VMware Fusion, or UTM).

---

## 📋 macOS Prerequisites

| Component | Minimum | Recommended | Notes |
|-----------|---------|-------------|-------|
| macOS | 12 Monterey | 14 Sonoma | Intel or Apple Silicon |
| Python | 3.9 | 3.12 | Via Homebrew or pyenv |
| PostgreSQL Client | 14 | 16 | For psql CLI tools |
| Excel for Mac | 2019 | Microsoft 365 | Python pipeline only |
| Homebrew | Any | Latest | Package manager |

---

## 🚀 Quick Install (Automated)

Run the provided install script:

```bash
git clone https://github.com/shiviyer/MinervaDB-XL-View.git
cd MinervaDB-XL-View
chmod +x install/install_macos.sh
./install/install_macos.sh
```

---

## 🔧 Manual Installation Steps

### Step 1: Install Homebrew

If you don't have Homebrew installed:

```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

For Apple Silicon (M1/M2/M3), add Homebrew to your PATH:

```bash
echo 'eval "$(/opt/homebrew/bin/brew shellenv)"' >> ~/.zprofile
eval "$(/opt/homebrew/bin/brew shellenv)"
```

Verify:
```bash
brew --version
```

---

### Step 2: Install Python 3.12

```bash
brew install python@3.12
```

Add to PATH (Apple Silicon):
```bash
echo 'export PATH="/opt/homebrew/opt/python@3.12/bin:$PATH"' >> ~/.zshrc
source ~/.zshrc
```

Add to PATH (Intel Mac):
```bash
echo 'export PATH="/usr/local/opt/python@3.12/bin:$PATH"' >> ~/.zshrc
source ~/.zshrc
```

Verify:
```bash
python3 --version   # Should print Python 3.12.x
pip3 --version
```

---

### Step 3: Install PostgreSQL Client Tools

```bash
brew install postgresql@16
```

Add to PATH (Apple Silicon):
```bash
echo 'export PATH="/opt/homebrew/opt/postgresql@16/bin:$PATH"' >> ~/.zshrc
source ~/.zshrc
```

Verify:
```bash
psql --version
```

---

### Step 4: Clone MinervaDB XL View

```bash
git clone https://github.com/shiviyer/MinervaDB-XL-View.git
cd MinervaDB-XL-View
```

---

### Step 5: Set Up Python Virtual Environment

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r python/requirements.txt
```

Verify all packages installed:
```bash
pip list | grep -E "psycopg2|pandas|openpyxl|sqlalchemy"
```

---

### Step 6: Configure Database Connection

```bash
cp config/config.example.ini config/config.ini
```

Edit the config file:
```bash
nano config/config.ini   # or: open -e config/config.ini
```

Update these values:
```ini
[database]
host     = your-postgresql-host
port     = 5432
database = your_database_name
username = excel_dashboard_ro
password = your_secure_password
ssl_mode = require

[output]
output_dir = ./output
```

> **Security:** Never commit config/config.ini to git. It is already in .gitignore.

---

### Step 7: Set Up Sample Database (Optional)

If you want to test with sample data:

```bash
# Create the schema
psql -h your-host -U postgres -d your_database \
     -f sql/setup/create_sample_schema.sql

# Seed sample data
psql -h your-host -U postgres -d your_database \
     -f sql/setup/seed_sample_data.sql
```

---

### Step 8: Test the Connection

```bash
source .venv/bin/activate
python3 python/pg_connector.py
```

Expected output:
```
[INFO] MinervaDB XL View: Testing connection...
[INFO] MinervaDB XL View: Connected to PostgreSQL 16.x
[INFO] MinervaDB XL View: Connection test passed!
```

---

### Step 9: Run the ETL Pipeline

```bash
source .venv/bin/activate

# Run all dashboards
python3 python/etl_pipeline.py --pipeline all --output ./output

# Run specific pipeline
python3 python/etl_pipeline.py --pipeline sales --output ./output
```

Output files will appear in the `./output/` directory as Excel (.xlsx) files, ready to open in Excel for Mac.

---

### Step 10: Open Dashboard in Excel for Mac

```bash
open output/MinervaDB_XL_View_*.xlsx
```

Or drag and drop the file into Excel for Mac.

---

## 🍎 macOS-Specific Workflow

Since Excel for Mac does not support VBA ODBC connections, use this workflow:

```
PostgreSQL DB
     │
     ▼
Python ETL Pipeline (python/etl_pipeline.py)
     │  • Extracts data via psycopg2
     │  • Transforms with pandas
     │  • Writes formatted .xlsx
     ▼
Excel for Mac
     │  • Open the generated .xlsx
     │  • Charts & pivot tables work natively
     │  • Refresh: re-run the ETL script
     ▼
Scheduled Auto-Refresh (cron / launchd)
```

---

## ⏰ Set Up Scheduled Auto-Refresh (launchd)

Create a launchd plist to refresh dashboards automatically:

```bash
cat > ~/Library/LaunchAgents/com.minervadb.xlview.refresh.plist << 'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.minervadb.xlview.refresh</string>
    <key>ProgramArguments</key>
    <array>
        <string>/bin/bash</string>
        <string>-c</string>
        <string>cd /path/to/MinervaDB-XL-View && source .venv/bin/activate && python3 python/etl_pipeline.py --pipeline all</string>
    </array>
    <key>StartInterval</key>
    <integer>900</integer>
    <key>RunAtLoad</key>
    <true/>
    <key>StandardOutPath</key>
    <string>/tmp/minervadb_xlview_refresh.log</string>
    <key>StandardErrorPath</key>
    <string>/tmp/minervadb_xlview_refresh_err.log</string>
</dict>
</plist>
EOF

# Load the scheduler
launchctl load ~/Library/LaunchAgents/com.minervadb.xlview.refresh.plist
```

> Replace `/path/to/MinervaDB-XL-View` with the actual path to your clone.

---

## 🪟 Full VBA Experience on macOS (via Windows VM)

For the complete Excel VBA + real-time ODBC experience, run Windows in a virtual machine:

### Option A: Parallels Desktop (Recommended)
1. Install Parallels Desktop (https://www.parallels.com/)
2. Create a Windows 11 VM
3. Install Microsoft 365 in the VM
4. Follow [docs/WINDOWS_INSTALL.md](WINDOWS_INSTALL.md) inside the VM
5. Mount your macOS MinervaDB-XL-View folder as a shared drive in Parallels

### Option B: VMware Fusion
1. Install VMware Fusion Pro (free for personal use)
2. Create a Windows 11 VM
3. Follow [docs/WINDOWS_INSTALL.md](WINDOWS_INSTALL.md) inside the VM

### Option C: UTM (Free, Apple Silicon)
1. Download UTM from https://mac.getutm.app/
2. Download Windows 11 ARM ISO (for M-series Macs)
3. Create VM and follow [docs/WINDOWS_INSTALL.md](WINDOWS_INSTALL.md)

---

## 🔑 Environment Variables (Recommended for Security)

Instead of storing the password in config.ini, use environment variables:

```bash
# Add to ~/.zshrc
export MINERVADB_HOST="your-postgresql-host"
export MINERVADB_PORT="5432"
export MINERVADB_DB="your_database"
export MINERVADB_USER="excel_dashboard_ro"
export MINERVADB_PASSWORD="your_secure_password"

source ~/.zshrc
```

Update config/config.ini to reference env vars:
```ini
[database]
host     = ${MINERVADB_HOST}
port     = ${MINERVADB_PORT}
database = ${MINERVADB_DB}
username = ${MINERVADB_USER}
password = ${MINERVADB_PASSWORD}
```

---

## 🧪 Run Tests

```bash
source .venv/bin/activate
pip install pytest
pytest tests/ -v
```

---

## ❓ Common macOS Issues

| Issue | Solution |
|-------|----------|
| `psycopg2` build fails on Apple Silicon | Use `pip install psycopg2-binary` instead |
| `brew: command not found` | Re-run Homebrew install and add to PATH |
| `python3: command not found` | Run `brew install python@3.12` and add to PATH |
| Permission denied on install_macos.sh | Run `chmod +x install/install_macos.sh` |
| Connection refused to PostgreSQL | Check host/port in config.ini and firewall rules |
| Excel won't open .xlsx from terminal | Run `open output/*.xlsx` or drag to Excel icon |

See [TROUBLESHOOTING.md](TROUBLESHOOTING.md) for more detailed solutions.

---

*MinervaDB XL View — Enterprise PostgreSQL Analytics Platform*
*macOS Installation Guide v1.0*
