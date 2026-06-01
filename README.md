# 🏛️ MinervaDB XL View

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PostgreSQL](https://img.shields.io/badge/PostgreSQL-15%2B-blue.svg)](https://www.postgresql.org/)
[![Excel VBA](https://img.shields.io/badge/Excel-VBA-green.svg)](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
[![Python](https://img.shields.io/badge/Python-3.9%2B-blue.svg)](https://www.python.org/)

> **MinervaDB XL View** — Enterprise-class Excel Dashboard & Analytics Platform connecting Microsoft Excel to PostgreSQL for real-time business intelligence, KPI tracking, and data-driven decision-making.

---

## 🚀 Overview

**MinervaDB XL View** is a production-ready framework for building enterprise-grade analytics dashboards in Microsoft Excel that connect directly to PostgreSQL databases. It includes VBA modules for native ODBC/OLEDB connectivity, Python-based ETL connectors, a curated SQL query library, and pre-built KPI dashboard templates.

**Key Features:**

- 🔌 **Direct PostgreSQL Connectivity** via ODBC, OLEDB, and Python (psycopg2/SQLAlchemy)
- 📊 **Enterprise KPI Dashboards** — Sales, Finance, HR, Operations, and Supply Chain
- 🔄 **Automated Data Refresh** — Scheduled and on-demand refresh with error handling
- 🔒 **Secure Credential Management** — Environment variables and encrypted config files
- 🏗️ **Modular VBA Architecture** — Reusable modules, class objects, and event-driven design
- 🐍 **Python ETL Bridge** — Pandas-powered data transformation pipeline
- 📋 **SQL Query Library** — 50+ optimized PostgreSQL queries for common analytics patterns
- 🎨 **Professional Formatting** — Conditional formatting, sparklines, and dynamic charts

---

## 🍎 Installation & Configuration — macOS (MacBook)

### System Requirements

| Component | Minimum | Recommended |
|-----------|---------|-------------|
| macOS | Ventura 13.0+ | Sonoma 14.0+ |
| Microsoft Excel | 2019 (16.x) | Microsoft 365 |
| Python | 3.9 | 3.11+ |
| PostgreSQL | 13 | 16+ |

> **Note:** Native Excel VBA macros on macOS have limited ODBC support compared to Windows. The **Python connector** is the recommended primary method on macOS.

### Step 1 — Install Homebrew (if not already installed)

```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

### Step 2 — Install PostgreSQL Client

```bash
brew install postgresql@16
echo 'export PATH="/opt/homebrew/opt/postgresql@16/bin:$PATH"' >> ~/.zshrc
source ~/.zshrc
```

### Step 3 — Install Python 3.11+

```bash
brew install python@3.11
# Verify installation
python3 --version
```

### Step 4 — Install psqlODBC Driver (for VBA connectivity)

```bash
brew install psqlodbc
# Or download the macOS package from https://odbc.postgresql.org/
```

### Step 5 — Clone the Repository

```bash
git clone https://github.com/shiviyer/MinervaDB-XL-View.git
cd MinervaDB-XL-View
```

### Step 6 — Set Up Python Virtual Environment

```bash
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r python/requirements.txt
```

### Step 7 — Configure Database Connection

```bash
cp config/config.example.ini config/config.ini
nano config/config.ini   # or use: open -e config/config.ini
```

Edit `config/config.ini` with your PostgreSQL credentials:

```ini
[postgresql]
host     = your-postgres-host
port     = 5432
database = your_database
username = excel_dashboard_ro
password = your_secure_password
sslmode  = require

[excel]
refresh_interval = 900   ; seconds (15 min)
max_rows         = 50000

[logging]
level   = INFO
logfile = logs/minervadb.log
```

### Step 8 — Set Up Environment Variables (Recommended)

```bash
echo 'export MINERVADB_HOST="your-postgres-host"' >> ~/.zshrc
echo 'export MINERVADB_PORT="5432"' >> ~/.zshrc
echo 'export MINERVADB_DB="your_database"' >> ~/.zshrc
echo 'export MINERVADB_USER="excel_dashboard_ro"' >> ~/.zshrc
echo 'export MINERVADB_PASSWORD="your_secure_password"' >> ~/.zshrc
source ~/.zshrc
```

### Step 9 — Initialize the Database

```bash
psql -h your-postgres-host -U postgres -f sql/setup/create_sample_schema.sql
psql -h your-postgres-host -U postgres -f sql/setup/seed_sample_data.sql
psql -h your-postgres-host -U postgres -f sql/setup/create_indexes.sql
```

### Step 10 — Test Python Connectivity

```bash
source venv/bin/activate
python python/pg_connector.py
# Expected: "MinervaDB XL View: Connection successful!"
```

### Step 11 — Import VBA Modules into Excel (macOS)

1. Open **Microsoft Excel** → Open one of the dashboard templates from `templates/`
2. Press **⌥ Option + F11** to open the VBA Editor (or use **Tools → Macros → Visual Basic Editor**)
3. In the VBA Editor: **File → Import File...**
4. Import all `.bas` and `.cls` files from the `vba/` directory in order:
   - `modErrorHandler.bas` (import first)
   - `modSecurity.bas`
   - `modConnection.bas`
   - `modQueryRunner.bas`
   - `modDataRefresh.bas`
   - `modChartBuilder.bas`
   - `modFormatting.bas`
   - `modExport.bas`
   - `clsPostgresConn.cls`
   - `clsDashboardSheet.cls`
   - `clsKPIWidget.cls`
5. Save the workbook as `.xlsm` (macro-enabled format)

### macOS Troubleshooting

| Issue | Solution |
|-------|----------|
| `psycopg2` install fails | Run: `pip install psycopg2-binary` |
| Excel VBA macros disabled | Go to **Excel → Preferences → Security** and enable macros |
| ODBC driver not found | Verify with: `odbcinst -j` and check `/usr/local/etc/odbcinst.ini` |
| SSL connection error | Add `sslmode = require` to config.ini and ensure server allows SSL |
| Permission denied on `logs/` | Run: `mkdir -p logs && chmod 755 logs` |

---

## 🪟 Installation & Configuration — Windows

### System Requirements

| Component | Minimum | Recommended |
|-----------|---------|-------------|
| Windows | Windows 10 (64-bit) | Windows 11 |
| Microsoft Excel | 2016 (32 or 64-bit) | Microsoft 365 (64-bit) |
| Python | 3.9 | 3.11+ |
| PostgreSQL | 13 | 16+ |

> **Important:** Match the bitness of your Excel installation with the psqlODBC driver (both 32-bit or both 64-bit).

### Step 1 — Install PostgreSQL (if not already installed)

1. Download the installer from [https://www.enterprisedb.com/downloads/postgres-postgresql-downloads](https://www.enterprisedb.com/downloads/postgres-postgresql-downloads)
2. Run the installer and select: **PostgreSQL Server**, **pgAdmin 4**, **Command Line Tools**
3. Note the installation path (default: `C:\Program Files\PostgreSQL\16`)
4. Add PostgreSQL to PATH:
   - Open **System Properties → Advanced → Environment Variables**
   - Under **System Variables**, edit **Path** and add: `C:\Program Files\PostgreSQL\16\bin`

### Step 2 — Install psqlODBC Driver

1. Download from [https://www.postgresql.org/ftp/odbc/versions/msi/](https://www.postgresql.org/ftp/odbc/versions/msi/)
2. Choose the version matching your Excel bitness:
   - **64-bit Excel** → `psqlodbc_16_xx_xxxx-x64.msi`
   - **32-bit Excel** → `psqlodbc_16_xx_xxxx-x86.msi`
3. Run the MSI installer with administrator privileges
4. Verify in **ODBC Data Sources** (search in Start Menu):
   - Open **ODBC Data Source Administrator (64-bit)**
   - Go to **Drivers** tab → confirm **PostgreSQL Unicode(x64)** is listed

### Step 3 — Install Python 3.11+

1. Download from [https://www.python.org/downloads/windows/](https://www.python.org/downloads/windows/)
2. Run installer — **check "Add Python to PATH"** before clicking Install
3. Verify installation:

```powershell
python --version
pip --version
```

### Step 4 — Install Git for Windows

1. Download from [https://git-scm.com/download/win](https://git-scm.com/download/win)
2. Use default options during installation
3. Verify: open **Git Bash** or **PowerShell** and run `git --version`

### Step 5 — Clone the Repository

```powershell
git clone https://github.com/shiviyer/MinervaDB-XL-View.git
cd MinervaDB-XL-View
```

### Step 6 — Set Up Python Virtual Environment

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install -r python\requirements.txt
```

> If you encounter a PowerShell execution policy error, run: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

### Step 7 — Configure Database Connection

```powershell
copy config\config.example.ini config\config.ini
notepad config\config.ini
```

Edit `config\config.ini` with your PostgreSQL credentials:

```ini
[postgresql]
host     = your-postgres-host
port     = 5432
database = your_database
username = excel_dashboard_ro
password = your_secure_password
sslmode  = require

[excel]
refresh_interval = 900   ; seconds (15 min)
max_rows         = 50000

[logging]
level   = INFO
logfile = logs\minervadb.log
```

### Step 8 — Set Up Environment Variables (Recommended)

```powershell
[System.Environment]::SetEnvironmentVariable("MINERVADB_HOST", "your-postgres-host", "User")
[System.Environment]::SetEnvironmentVariable("MINERVADB_PORT", "5432", "User")
[System.Environment]::SetEnvironmentVariable("MINERVADB_DB", "your_database", "User")
[System.Environment]::SetEnvironmentVariable("MINERVADB_USER", "excel_dashboard_ro", "User")
[System.Environment]::SetEnvironmentVariable("MINERVADB_PASSWORD", "your_secure_password", "User")
```

Or set them via GUI: **System Properties → Advanced → Environment Variables → New (User variables)**

### Step 9 — Configure ODBC Data Source

1. Open **ODBC Data Source Administrator (64-bit)** from Start Menu
2. Click **Add...** on the **System DSN** tab
3. Select **PostgreSQL Unicode(x64)** → click **Finish**
4. Fill in the configuration:

```
Data Source:  MinervaDB_XL
Description:  MinervaDB XL View PostgreSQL Connection
Database:     your_database
Server:       your-postgres-host
Port:         5432
User Name:    excel_dashboard_ro
Password:     your_secure_password
SSL Mode:     require
```

5. Click **Test** to verify the connection → click **Save**

### Step 10 — Initialize the Database

```powershell
psql -h your-postgres-host -U postgres -f sql\setup\create_sample_schema.sql
psql -h your-postgres-host -U postgres -f sql\setup\seed_sample_data.sql
psql -h your-postgres-host -U postgres -f sql\setup\create_indexes.sql
```

### Step 11 — Test Python Connectivity

```powershell
.\venv\Scripts\Activate.ps1
python python\pg_connector.py
# Expected: "MinervaDB XL View: Connection successful!"
```

### Step 12 — Import VBA Modules into Excel (Windows)

1. Open **Microsoft Excel** → Open one of the dashboard templates from `templates\`
2. Press **Alt + F11** to open the VBA Editor
3. In the Project Explorer: right-click your workbook → **Import File...**
4. Import all `.bas` and `.cls` files from the `vba\` directory in order:
   - `modErrorHandler.bas` (import first)
   - `modSecurity.bas`
   - `modConnection.bas`
   - `modQueryRunner.bas`
   - `modDataRefresh.bas`
   - `modChartBuilder.bas`
   - `modFormatting.bas`
   - `modExport.bas`
   - `clsPostgresConn.cls`
   - `clsDashboardSheet.cls`
   - `clsKPIWidget.cls`
5. In the VBA Editor, go to **Tools → References** and enable:
   - **Microsoft ActiveX Data Objects 6.1 Library**
   - **Microsoft Scripting Runtime**
6. Save the workbook as `.xlsm` (macro-enabled format)
7. When prompted, **Enable Macros** on open

### Step 13 — Connect Excel to PostgreSQL via ODBC (Windows VBA)

Update the connection string in `modConnection.bas` to use your DSN:

```vba
' MinervaDB XL View — Windows ODBC Connection
Const CONN_STRING = "DSN=MinervaDB_XL;UID=excel_dashboard_ro;PWD=your_password;"

Sub ConnectToMinervaDB()
    Dim conn As New ADODB.Connection
    conn.Open CONN_STRING
    If conn.State = adStateOpen Then
        MsgBox "MinervaDB XL View: Connected!", vbInformation
        conn.Close
    End If
End Sub
```

### Windows Troubleshooting

| Issue | Solution |
|-------|----------|
| ODBC driver architecture mismatch | Ensure Excel and psqlODBC are both 64-bit or both 32-bit |
| `Run-time error '3706'` in VBA | Check ODBC DSN name matches `CONN_STRING` exactly |
| `psycopg2` install fails | Run: `pip install psycopg2-binary` |
| Macros blocked by security | Go to **File → Options → Trust Center → Trust Center Settings → Macro Settings** → Enable macros |
| PowerShell script blocked | Run: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` |
| SSL handshake failure | Download root CA cert and add `sslrootcert=path\to\root.crt` to ODBC DSN |
| `logs\` directory missing | Run: `mkdir logs` in the project root |

---

## 📁 Repository Structure

```
MinervaDB-XL-View/
├── config/
│   └── config.example.ini        # Configuration template
├── docs/
│   ├── QUICK_START.md             # 15-minute setup guide
│   ├── INSTALLATION.md            # Detailed installation guide
│   ├── VBA_API_REFERENCE.md       # VBA module API docs
│   ├── PYTHON_GUIDE.md            # Python connector guide
│   ├── SQL_QUERY_GUIDE.md         # SQL query library reference
│   ├── DASHBOARD_GUIDE.md         # Dashboard creation guide
│   └── TROUBLESHOOTING.md         # Common issues and solutions
├── python/
│   ├── pg_connector.py            # Core PostgreSQL connector
│   ├── etl_pipeline.py            # ETL data transformation pipeline
│   ├── excel_writer.py            # Excel file writer utility
│   ├── scheduler.py               # Automated refresh scheduler
│   ├── data_validator.py          # Data quality validation
│   ├── config_manager.py          # Secure configuration management
│   └── requirements.txt           # Python dependencies
├── sql/
│   ├── kpi/
│   │   ├── sales_kpis.sql         # Sales performance KPIs
│   │   ├── finance_kpis.sql       # Financial metrics KPIs
│   │   ├── hr_kpis.sql            # HR analytics KPIs
│   │   ├── operations_kpis.sql    # Operational KPIs
│   │   └── supply_chain_kpis.sql  # Supply chain KPIs
│   ├── reports/
│   │   ├── monthly_summary.sql    # Monthly summary report
│   │   └── trend_analysis.sql     # Trend analysis queries
│   ├── views/
│   │   └── vw_executive_kpis.sql  # Executive KPI view
│   └── setup/
│       ├── create_sample_schema.sql  # Sample schema
│       ├── seed_sample_data.sql      # Sample data
│       └── create_indexes.sql        # Performance indexes
├── vba/
│   ├── modConnection.bas          # PostgreSQL connection manager
│   ├── modQueryRunner.bas         # SQL query execution engine
│   ├── modDataRefresh.bas         # Automated data refresh
│   ├── modChartBuilder.bas        # Dynamic chart builder
│   ├── modFormatting.bas          # Professional formatting
│   ├── modErrorHandler.bas        # Centralized error handling
│   ├── modSecurity.bas            # Credential security
│   ├── modExport.bas              # Export to PDF/CSV
│   ├── clsPostgresConn.cls        # PostgreSQL connection class
│   ├── clsDashboardSheet.cls      # Dashboard sheet class
│   └── clsKPIWidget.cls           # KPI widget class
├── templates/
│   ├── sales_dashboard.xlsx       # Sales dashboard template
│   ├── finance_dashboard.xlsx     # Finance dashboard template
│   └── executive_summary.xlsx     # Executive summary template
├── tests/
│   ├── test_connection.py         # Connection tests
│   ├── test_etl_pipeline.py       # ETL pipeline tests
│   └── test_data_validator.py     # Data validation tests
├── .gitignore
├── LICENSE
└── README.md
```

---

## ⚡ Quick Start

### Prerequisites
- Microsoft Excel 2016+ (Microsoft 365 recommended)
- PostgreSQL 13+
- Python 3.9+
- PostgreSQL ODBC Driver (psqlODBC)

### 1. Clone and Configure

```bash
git clone https://github.com/shiviyer/MinervaDB-XL-View.git
cd MinervaDB-XL-View
cp config/config.example.ini config/config.ini
# Edit config/config.ini with your PostgreSQL credentials
```

### 2. Python Setup

```bash
pip install -r python/requirements.txt
python python/pg_connector.py  # Test connection
```

### 3. Database Setup

```bash
psql -U postgres -f sql/setup/create_sample_schema.sql
psql -U postgres -f sql/setup/seed_sample_data.sql
```

### 4. Excel VBA Import

Open Excel → Developer → Visual Basic → File → Import File → select all `.bas` files from `vba/`

---

## 🔌 VBA Connection Example

```vba
' MinervaDB XL View - Quick Connection
Sub ConnectToMinervaDB()
    Dim conn As New clsPostgresConn
    conn.Host     = "your-postgres-host"
    conn.Port     = 5432
    conn.Database = "your_database"
    conn.Username = "excel_dashboard_ro"
    conn.Password = GetSecurePassword()

    If conn.Connect() Then
        MsgBox "MinervaDB XL View: Connected!", vbInformation
        Call LoadSalesDashboard(conn)
        conn.Disconnect
    End If
End Sub
```

---

## 🐍 Python Connection Example

```python
from python.pg_connector import MinervaDBConnector

connector = MinervaDBConnector(config_path='config/config.ini')
with connector.get_connection() as conn:
    df = connector.execute_query(conn,
        "SELECT * FROM vw_sales_dashboard "
        "WHERE report_date >= CURRENT_DATE - INTERVAL '30 days'"
    )
    connector.to_excel(df, 'Sales_Dashboard', 'output/sales_report.xlsx')
```

---

## 📊 Dashboard Templates

| Dashboard | Metrics | Refresh Rate |
|-----------|---------|--------------|
| Sales Performance | Revenue, Pipeline, Win Rate, ACV | 15 min |
| Financial Overview | P&L, Cash Flow, Budget vs Actual | 1 hour |
| HR Analytics | Headcount, Attrition, Performance | Daily |
| Operations | Throughput, Quality, Lead Time | 30 min |
| Executive Summary | All KPIs, Trends, Forecasts | 1 hour |
| Supply Chain | Inventory, Supplier, Fulfillment | 15 min |

---

## 🔒 Security

MinervaDB XL View implements enterprise security standards:

- Read-only PostgreSQL role (`excel_dashboard_ro`) for dashboard access
- AES-256 encryption for stored credentials
- Environment variable support (no hardcoded passwords)
- SSL/TLS encrypted connections
- Parameterized queries (SQL injection prevention)
- Audit logging for all data access

---

## 📄 License

MIT License — see [LICENSE](LICENSE) for details.

---

## 🏛️ About MinervaDB XL View

MinervaDB XL View bridges the gap between enterprise PostgreSQL databases and business users who rely on Microsoft Excel for analytics. Built for data engineers, analysts, and BI teams who need enterprise-grade connectivity without sacrificing Excel's flexibility.

**Powered by MinervaDB** | Enterprise PostgreSQL Analytics Platform
