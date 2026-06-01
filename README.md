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

## 📁 Repository Structure

```
MinervaDB-XL-View/
├── config/
│   └── config.example.ini          # Configuration template
├── docs/
│   ├── QUICK_START.md              # 15-minute setup guide
│   ├── INSTALLATION.md             # Detailed installation guide
│   ├── VBA_API_REFERENCE.md        # VBA module API docs
│   ├── PYTHON_GUIDE.md             # Python connector guide
│   ├── SQL_QUERY_GUIDE.md          # SQL query library reference
│   ├── DASHBOARD_GUIDE.md          # Dashboard creation guide
│   └── TROUBLESHOOTING.md          # Common issues and solutions
├── python/
│   ├── pg_connector.py             # Core PostgreSQL connector
│   ├── etl_pipeline.py             # ETL data transformation pipeline
│   ├── excel_writer.py             # Excel file writer utility
│   ├── scheduler.py                # Automated refresh scheduler
│   ├── data_validator.py           # Data quality validation
│   ├── config_manager.py           # Secure configuration management
│   └── requirements.txt            # Python dependencies
├── sql/
│   ├── kpi/
│   │   ├── sales_kpis.sql          # Sales performance KPIs
│   │   ├── finance_kpis.sql        # Financial metrics KPIs
│   │   ├── hr_kpis.sql             # HR analytics KPIs
│   │   ├── operations_kpis.sql     # Operational KPIs
│   │   └── supply_chain_kpis.sql   # Supply chain KPIs
│   ├── reports/
│   │   ├── monthly_summary.sql     # Monthly summary report
│   │   └── trend_analysis.sql      # Trend analysis queries
│   ├── views/
│   │   └── vw_executive_kpis.sql   # Executive KPI view
│   └── setup/
│       ├── create_sample_schema.sql # Sample schema
│       ├── seed_sample_data.sql     # Sample data
│       └── create_indexes.sql       # Performance indexes
├── vba/
│   ├── modConnection.bas           # PostgreSQL connection manager
│   ├── modQueryRunner.bas          # SQL query execution engine
│   ├── modDataRefresh.bas          # Automated data refresh
│   ├── modChartBuilder.bas         # Dynamic chart builder
│   ├── modFormatting.bas           # Professional formatting
│   ├── modErrorHandler.bas         # Centralized error handling
│   ├── modSecurity.bas             # Credential security
│   ├── modExport.bas               # Export to PDF/CSV
│   ├── clsPostgresConn.cls         # PostgreSQL connection class
│   ├── clsDashboardSheet.cls       # Dashboard sheet class
│   └── clsKPIWidget.cls            # KPI widget class
├── templates/
│   ├── sales_dashboard.xlsx        # Sales dashboard template
│   ├── finance_dashboard.xlsx      # Finance dashboard template
│   └── executive_summary.xlsx      # Executive summary template
├── tests/
│   ├── test_connection.py          # Connection tests
│   ├── test_etl_pipeline.py        # ETL pipeline tests
│   └── test_data_validator.py      # Data validation tests
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
    conn.Host = "your-postgres-host"
    conn.Port = 5432
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
|-----------|---------|-------------|
| **Sales Performance** | Revenue, Pipeline, Win Rate, ACV | 15 min |
| **Financial Overview** | P&L, Cash Flow, Budget vs Actual | 1 hour |
| **HR Analytics** | Headcount, Attrition, Performance | Daily |
| **Operations** | Throughput, Quality, Lead Time | 30 min |
| **Executive Summary** | All KPIs, Trends, Forecasts | 1 hour |
| **Supply Chain** | Inventory, Supplier, Fulfillment | 15 min |

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

**MinervaDB XL View** bridges the gap between enterprise PostgreSQL databases and business users who rely on Microsoft Excel for analytics. Built for data engineers, analysts, and BI teams who need enterprise-grade connectivity without sacrificing Excel's flexibility.

*Powered by MinervaDB | Enterprise PostgreSQL Analytics Platform*
