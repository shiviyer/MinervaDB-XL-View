# Quick Start Guide

## Excel-PostgreSQL-Dashboard — Up and Running in 15 Minutes

This guide will have you connected to PostgreSQL and viewing live data in Excel within 15 minutes.

---

## Prerequisites Checklist

Before starting, ensure you have:

- [ ] Microsoft Excel 2016 or later (Excel 365 recommended)
- [ ] PostgreSQL 13+ server access
- [ ] Python 3.9+ installed
- [ ] Admin access to install the PostgreSQL ODBC driver

---

## Step 1: Install PostgreSQL ODBC Driver (5 min)

1. Download from [https://www.postgresql.org/ftp/odbc/versions/](https://www.postgresql.org/ftp/odbc/versions/)
2. Choose `psqlodbc_16_x.msi` (latest version)
3. Run the installer with default settings
4. Restart Excel after installation

**Verify installation:** Open Windows ODBC Data Source Administrator (64-bit) and check that "PostgreSQL Unicode(x64)" appears in the driver list.

---

## Step 2: Clone & Install (2 min)

```bash
# Clone the repository
git clone https://github.com/shiviyer/Excel-PostgreSQL-Dashboard.git
cd Excel-PostgreSQL-Dashboard

# Install Python dependencies
pip install -r python/requirements.txt

# Verify installation
python -c "import psycopg2; import pandas; print('Dependencies OK')"
```

---

## Step 3: Configure Connection (2 min)

```bash
# Copy the example config file
cp config/config.example.ini config/config.ini
```

Edit `config/config.ini` with your PostgreSQL details:

```ini
[postgresql]
host     = your-postgres-host
port     = 5432
database = your_database
username = excel_dashboard_ro
password = your_password
schema   = public
sslmode  = prefer
```

**Test the connection:**
```bash
python python/pg_connector.py config/config.ini
```

Expected output:
```
✓ Connected: your_database@your-host:5432
PostgreSQL Version: PostgreSQL 16.x on ...
Schemas: ['public', 'information_schema']
```

---

## Step 4: Set Up Sample Data (Optional, 2 min)

If you want to use the sample dashboard templates immediately:

```bash
# Create schema and sample tables
psql -U your_username -d your_database -f sql/setup/create_sample_schema.sql

# Load sample data (if available)
psql -U your_username -d your_database -f sql/setup/seed_sample_data.sql
```

---

## Step 5: Configure Excel VBA (4 min)

### 5a. Enable Developer Tab
1. File → Options → Customize Ribbon
2. Check "Developer" in the right panel
3. Click OK

### 5b. Import VBA Modules
1. Open Excel and press `Alt + F11` to open the VBA IDE
2. In the Project Explorer, right-click your workbook name
3. Select **Import File**
4. Import all `.bas` files from the `vba/` folder in this order:
   - `modErrorHandler.bas` (import first)
      - `modConnection.bas`
         - `modQueryRunner.bas`
            - `modDataRefresh.bas`
               - `modFormatting.bas`
                  - `modChartBuilder.bas`
                     - `modSecurity.bas`
                        - `modExport.bas`
                        5. Import `.cls` files:
                           - `clsPostgresConn.cls`
                              - `clsDashboardSheet.cls`
                                 - `clsKPIWidget.cls`
                                 6. Save the workbook as `.xlsm` (Excel Macro-Enabled Workbook)

                                 ### 5c. Set Macro Security
                                 1. File → Options → Trust Center → Trust Center Settings
                                 2. Macro Settings → Select "Enable all macros" (or "Disable with notification")
                                 3. Click OK

                                 ---

                                 ## Step 6: Connect and Refresh Data

                                 ### Using VBA (in-Excel)
                                 Press `Alt + F8` → Select `ShowConnectionDialog` → Run

                                 Fill in your connection details when prompted.

                                 ### Using Python (command line)
                                 ```python
                                 from python.pg_connector import connect

                                 with connect("config/config.ini") as conn:
                                     # Get sales KPIs
                                         df = conn.query_from_file("sql/kpi/sales_kpis.sql")
                                             print(df.head())
                                             ```

                                             ---

                                             ## Step 7: Open a Dashboard Template

                                             1. Open any template from the `templates/` folder
                                             2. Enable macros when prompted
                                             3. Click the **Refresh Data** button on the dashboard
                                             4. Enter your connection credentials
                                             5. Watch your dashboard populate with live PostgreSQL data!

                                             ---

                                             ## Troubleshooting Quick Fixes

                                             | Problem | Solution |
                                             |---------|----------|
                                             | "Driver not found" error | Reinstall psqlODBC, restart Excel |
                                             | "Authentication failed" | Check username/password in config.ini |
                                             | "Connection timeout" | Check firewall, VPN, PostgreSQL `pg_hba.conf` |
                                             | Macros won't run | Enable macros in Trust Center settings |
                                             | Blank dashboard | Run `ShowConnectionDialog` first |
                                             | Python import error | Run `pip install -r python/requirements.txt` |

                                             ---

                                             ## Next Steps

                                             - Read `docs/VBA_API_REFERENCE.md` for full VBA module documentation
                                             - Read `docs/SQL_QUERY_GUIDE.md` to customize queries for your schema
                                             - Read `docs/DASHBOARD_GUIDE.md` to create custom dashboards
                                             - Set up automated refresh: see `docs/PYTHON_GUIDE.md`

                                             ---

                                             *For enterprise support: [MinervaDB](https://minervadb.com) | Author: Shiv Iyer*
