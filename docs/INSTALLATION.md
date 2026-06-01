# MinervaDB XL View - Installation Guide

Complete step-by-step installation guide for MinervaDB XL View.

---

## System Requirements

| Component | Minimum | Recommended |
|-----------|---------|-------------|
| OS | Windows 10 / macOS 12 / Ubuntu 20.04 | Windows 11 / macOS 14 / Ubuntu 22.04 |
| Excel | Excel 2016 | Microsoft 365 |
| PostgreSQL | 13.x | 15.x or 16.x |
| Python | 3.9 | 3.11 or 3.12 |
| RAM | 8 GB | 16 GB |

---

## Step 1: Clone the Repository

```bash
git clone https://github.com/shiviyer/MinervaDB-XL-View.git
cd MinervaDB-XL-View
```

---

## Step 2: Install PostgreSQL ODBC Driver

**Windows:** Download psqlODBC from https://www.postgresql.org/ftp/odbc/versions/msi/ and install the 64-bit version.

**macOS:**
```bash
brew install psqlodbc
```

**Ubuntu/Debian:**
```bash
sudo apt-get install odbc-postgresql unixodbc-dev
```

---

## Step 3: Configure the Connection

```bash
cp config/config.example.ini config/config.ini
# Edit config/config.ini with your PostgreSQL credentials
```

> Security: Never commit config/config.ini. Store passwords in environment variables.

---

## Step 4: Set Up the Database

```bash
psql -h your-host -U postgres -d your_database -f sql/setup/create_sample_schema.sql
psql -h your-host -U postgres -d your_database -f sql/setup/seed_sample_data.sql
```

---

## Step 5: Set Up Python Environment

```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r python/requirements.txt
python python/pg_connector.py  # Test connection
```

---

## Step 6: Import VBA Modules into Excel

1. Open Excel and press **Alt+F11** to open VBA Editor
2. Go to **File > Import File**
3. Import all .bas files from the vba/ directory:
   - modErrorHandler.bas (first)
   - modConnection.bas
   - modQueryRunner.bas
   - modDataRefresh.bas
4. Save workbook as .xlsm (macro-enabled format)

---

## Step 7: Test the Connection

In VBA Immediate Window (Ctrl+G):
```vba
Call modConnection.TestConnection()
```

---

## Step 8: Run Initial Data Load

```bash
python python/etl_pipeline.py --output output
```

---

## Verification Checklist

- [ ] Repository cloned
- [ ] ODBC driver installed
- [ ] config/config.ini configured
- [ ] Database schema created
- [ ] Python environment active with packages installed
- [ ] python/pg_connector.py returns success
- [ ] VBA modules imported into Excel
- [ ] Excel saved as .xlsm
- [ ] VBA connection test passes

---

## Troubleshooting

| Problem | Cause | Solution |
|---------|-------|----------|
| ODBC Driver Not Found | Driver not installed | Run psqlODBC installer |
| Connection Refused | Wrong host/port | Check config.ini and firewall |
| Authentication Failed | Wrong credentials | Verify in PostgreSQL pg_hba.conf |
| VBA Compile Error | Import order wrong | Import modErrorHandler.bas first |
| Python ImportError | Packages missing | pip install -r requirements.txt |

See [TROUBLESHOOTING.md](TROUBLESHOOTING.md) for detailed solutions.

---

## Security Best Practices

1. Use dedicated read-only PostgreSQL role (excel_dashboard_ro)
2. Enable SSL/TLS (ssl_mode = require)
3. Store passwords in environment variables
4. Restrict network access to PostgreSQL port 5432
5. Rotate credentials regularly

---

*MinervaDB XL View - Enterprise PostgreSQL Analytics Platform*
