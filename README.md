# 📊 Excel-PostgreSQL-Dashboard

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PostgreSQL](https://img.shields.io/badge/PostgreSQL-15%2B-blue.svg)](https://www.postgresql.org/)
[![Excel VBA](https://img.shields.io/badge/Excel-VBA-green.svg)](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
[![Python](https://img.shields.io/badge/Python-3.9%2B-blue.svg)](https://www.python.org/)

> **Enterprise-class Excel Dashboard & Analytics Repository** connecting Microsoft Excel to PostgreSQL for real-time business intelligence, KPI tracking, and data-driven decision-making.
>
> ---
>
> ## 🚀 Overview
>
> The **Excel-PostgreSQL-Dashboard** repository provides a complete, production-ready framework for building enterprise-grade analytics dashboards in Microsoft Excel that connect directly to PostgreSQL databases. It includes VBA modules for native ODBC/OLEDB connectivity, Python-based ETL connectors, a curated SQL query library, and pre-built KPI dashboard templates.
>
> **Key Features:**
> - 🔌 **Direct PostgreSQL Connectivity** via ODBC, OLEDB, and Python (psycopg2/SQLAlchemy)
> - - 📈 **Enterprise KPI Dashboards** — Sales, Finance, HR, Operations, and Supply Chain
>   - - 🔄 **Automated Data Refresh** — Scheduled and on-demand refresh with error handling
>     - - 🔒 **Secure Credential Management** — Environment variables & encrypted config files
>       - - 📐 **Modular VBA Architecture** — Reusable modules, class objects, and event-driven design
>         - - 🐍 **Python ETL Bridge** — Pandas-powered data transformation pipeline
>           - - 📋 **SQL Query Library** — 50+ optimized PostgreSQL queries for common analytics patterns
>             - - 🎨 **Professional Formatting** — Conditional formatting, sparklines, and dynamic charts
>              
>               - ---
>
> ## 📁 Repository Structure
>
> ```
> Excel-PostgreSQL-Dashboard/
> │
> ├── 📂 vba/                          # Excel VBA Modules
> │   ├── modConnection.bas            # PostgreSQL connection manager
> │   ├── modQueryRunner.bas           # SQL query execution engine
> │   ├── modDataRefresh.bas           # Scheduled data refresh module
> │   ├── modChartBuilder.bas          # Dynamic chart generation
> │   ├── modFormatting.bas            # Dashboard formatting utilities
> │   ├── modErrorHandler.bas          # Centralized error handling
> │   ├── modSecurity.bas              # Credential encryption module
> │   ├── modExport.bas                # PDF/Excel export utilities
> │   ├── clsPostgresConn.cls          # PostgreSQL connection class
> │   ├── clsDashboardSheet.cls        # Dashboard sheet manager class
> │   └── clsKPIWidget.cls             # KPI widget class
> │
> ├── 📂 python/                       # Python ETL & Connector Scripts
> │   ├── pg_connector.py              # Core PostgreSQL connector
> │   ├── etl_pipeline.py              # ETL data transformation pipeline
> │   ├── excel_writer.py              # Excel file writer (openpyxl)
> │   ├── scheduler.py                 # Automated refresh scheduler
> │   ├── data_validator.py            # Data quality validation
> │   ├── config_manager.py            # Secure configuration manager
> │   └── requirements.txt             # Python dependencies
> │
> ├── 📂 sql/                          # PostgreSQL SQL Query Library
> │   ├── 📂 kpi/                      # KPI metric queries
> │   │   ├── sales_kpis.sql           # Sales performance metrics
> │   │   ├── finance_kpis.sql         # Financial KPI queries
> │   │   ├── hr_kpis.sql              # HR & workforce analytics
> │   │   ├── operations_kpis.sql      # Operational efficiency metrics
> │   │   └── supply_chain_kpis.sql    # Supply chain analytics
> │   ├── 📂 reports/                  # Report-level queries
> │   │   ├── monthly_summary.sql      # Monthly executive summary
> │   │   ├── trend_analysis.sql       # Time-series trend analysis
> │   │   ├── cohort_analysis.sql      # Customer cohort analysis
> │   │   └── variance_report.sql      # Budget vs actual variance
> │   ├── 📂 views/                    # PostgreSQL view definitions
> │   │   ├── vw_sales_dashboard.sql   # Sales dashboard view
> │   │   ├── vw_finance_summary.sql   # Finance summary view
> │   │   └── vw_executive_kpis.sql    # Executive KPI view
> │   └── 📂 setup/                    # Database setup scripts
> │       ├── create_sample_schema.sql # Sample schema creation
> │       ├── seed_sample_data.sql     # Sample data seeding
> │       └── create_indexes.sql       # Performance indexes
> │
> ├── 📂 templates/                    # Excel Dashboard Templates
> │   ├── Executive_Dashboard.xlsx     # C-Suite executive dashboard
> │   ├── Sales_Analytics.xlsx         # Sales performance dashboard
> │   ├── Finance_Dashboard.xlsx       # Finance & P&L dashboard
> │   ├── HR_Analytics.xlsx            # HR workforce dashboard
> │   ├── Operations_Dashboard.xlsx    # Operations KPI dashboard
> │   └── Supply_Chain.xlsx            # Supply chain analytics
> │
> ├── 📂 config/                       # Configuration Files
> │   ├── config.example.ini           # Example configuration file
> │   ├── dsn_setup_guide.md           # ODBC DSN setup guide
> │   └── connection_strings.md        # Connection string reference
> │
> ├── 📂 docs/                         # Documentation
> │   ├── INSTALLATION.md              # Installation guide
> │   ├── QUICK_START.md               # Quick start guide
> │   ├── VBA_API_REFERENCE.md         # VBA module API reference
> │   ├── PYTHON_GUIDE.md              # Python connector guide
> │   ├── SQL_QUERY_GUIDE.md           # SQL query library guide
> │   ├── DASHBOARD_GUIDE.md           # Dashboard template guide
> │   ├── TROUBLESHOOTING.md           # Troubleshooting guide
> │   └── ARCHITECTURE.md              # System architecture overview
> │
> ├── 📂 tests/                        # Test Scripts
> │   ├── test_connection.py           # Connection test suite
> │   ├── test_etl_pipeline.py         # ETL pipeline tests
> │   └── test_data_validator.py       # Data validation tests
> │
> ├── .gitignore                       # Git ignore rules
> ├── README.md                        # This file
> └── LICENSE                          # MIT License
> ```
>
> ---
>
> ## ⚙️ Prerequisites
>
> ### System Requirements
> | Component | Minimum Version |
> |-----------|----------------|
> | Microsoft Excel | 2016 or later (365 recommended) |
> | PostgreSQL | 13+ |
> | Python | 3.9+ |
> | Windows OS | Windows 10/11 or Windows Server 2019+ |
> | ODBC Driver | PostgreSQL ODBC Driver (psqlODBC) 13.02+ |
>
> ### Required Software
> 1. **PostgreSQL ODBC Driver** — [Download psqlODBC](https://www.postgresql.org/ftp/odbc/versions/)
> 2. 2. **Python 3.9+** — [Download Python](https://www.python.org/downloads/)
>    3. 3. **Microsoft Excel 2016+** with Developer tab enabled
>      
>       4. ---
>      
>       5. ## 🛠️ Installation
>      
>       6. ### Step 1: Clone the Repository
> ```bash
> git clone https://github.com/shiviyer/Excel-PostgreSQL-Dashboard.git
> cd Excel-PostgreSQL-Dashboard
> ```
>
> ### Step 2: Install Python Dependencies
> ```bash
> pip install -r python/requirements.txt
> ```
>
> ### Step 3: Configure PostgreSQL ODBC DSN
> Follow the guide in `config/dsn_setup_guide.md` to set up your ODBC Data Source Name.
>
> ### Step 4: Configure Connection Settings
> ```ini
> # config/config.ini (copy from config.example.ini)
> [postgresql]
> host     = your-postgres-host
> port     = 5432
> database = your_database
> username = your_username
> schema   = public
> ```
>
> ### Step 5: Set Up Sample Database (Optional)
> ```bash
> psql -U your_username -d your_database -f sql/setup/create_sample_schema.sql
> psql -U your_username -d your_database -f sql/setup/seed_sample_data.sql
> psql -U your_username -d your_database -f sql/setup/create_indexes.sql
> ```
>
> ### Step 6: Import VBA Modules into Excel
> 1. Open Excel → Press `Alt + F11` to open VBA IDE
> 2. 2. Right-click on your workbook → Import File
>    3. 3. Import all `.bas` and `.cls` files from the `vba/` directory
>       4. 4. Save as `.xlsm` (macro-enabled workbook)
>         
>          5. ---
>         
>          6. ## 🔌 PostgreSQL Connection (VBA)
>         
>          7. ### Basic Connection Example
> ```vba
> ' Initialize connection using the clsPostgresConn class
> Dim pgConn As New clsPostgresConn
> pgConn.Host     = "localhost"
> pgConn.Port     = 5432
> pgConn.Database = "your_database"
> pgConn.Username = "your_username"
>
> If pgConn.Connect() Then
>     MsgBox "Connected to PostgreSQL successfully!"
>     pgConn.Disconnect
> End If
> ```
>
> ### Execute a Query and Load to Sheet
> ```vba
> ' Run a query and populate a worksheet
> Dim qRunner As New clsQueryRunner
> qRunner.Connection = pgConn
> qRunner.SQL = "SELECT * FROM vw_sales_dashboard WHERE period = CURRENT_MONTH"
> qRunner.LoadToSheet ActiveSheet, startRow:=2, startCol:=1
> ```
>
> ---
>
> ## 🐍 Python Connector
>
> ### Quick Connect
> ```python
> from python.pg_connector import PostgreSQLConnector
>
> conn = PostgreSQLConnector.from_config("config/config.ini")
> df = conn.query_to_dataframe("SELECT * FROM vw_executive_kpis")
> print(df.head())
> ```
>
> ### ETL Pipeline
> ```python
> from python.etl_pipeline import ETLPipeline
>
> pipeline = ETLPipeline(config_path="config/config.ini")
> pipeline.extract(query_file="sql/kpi/sales_kpis.sql")
> pipeline.transform()
> pipeline.load_to_excel("templates/Sales_Analytics.xlsx", sheet="Data")
> ```
>
> ---
>
> ## 📊 Dashboard Templates
>
> | Template | Description | KPIs Included |
> |----------|-------------|---------------|
> | **Executive Dashboard** | C-Suite summary view | Revenue, EBITDA, Customer NPS, Employee Count |
> | **Sales Analytics** | Sales team performance | Pipeline, Win Rate, ARR, Quota Attainment |
> | **Finance Dashboard** | P&L and cash flow | Revenue, COGS, Gross Margin, Cash Position |
> | **HR Analytics** | Workforce metrics | Headcount, Attrition, Time-to-Hire, Engagement |
> | **Operations Dashboard** | Operational efficiency | SLA Compliance, Ticket Volume, Uptime, MTTR |
> | **Supply Chain** | Logistics and inventory | Inventory Turns, Lead Time, Fill Rate, OTD% |
>
> ---
>
> ## 🗄️ SQL Query Library
>
> ### Example: Sales KPI Query
> ```sql
> -- Monthly Sales Performance with YoY Comparison
> SELECT
>     DATE_TRUNC('month', order_date)     AS month,
>     SUM(revenue)                         AS total_revenue,
>     COUNT(DISTINCT customer_id)          AS unique_customers,
>     SUM(revenue) / NULLIF(COUNT(*), 0)  AS avg_order_value,
>     SUM(revenue) - LAG(SUM(revenue)) OVER (ORDER BY DATE_TRUNC('month', order_date))
>                                          AS mom_change
> FROM orders
> WHERE order_date >= NOW() - INTERVAL '13 months'
> GROUP BY 1
> ORDER BY 1;
> ```
>
> ### Example: HR Attrition Query
> ```sql
> -- Rolling 12-Month Attrition Rate
> SELECT
>     dept_name,
>     COUNT(*) FILTER (WHERE status = 'Active')       AS active_headcount,
>     COUNT(*) FILTER (WHERE termination_date IS NOT NULL
>                      AND termination_date >= NOW() - INTERVAL '12 months')
>                                                      AS terminations_12m,
>     ROUND(
>         COUNT(*) FILTER (WHERE termination_date IS NOT NULL
>                          AND termination_date >= NOW() - INTERVAL '12 months')::NUMERIC
>         / NULLIF(COUNT(*), 0) * 100, 2
>     )                                                AS attrition_rate_pct
> FROM employees
> JOIN departments USING (dept_id)
> GROUP BY dept_name
> ORDER BY attrition_rate_pct DESC;
> ```
>
> ---
>
> ## 🔒 Security Best Practices
>
> - **Never commit credentials** — Use `config/config.ini` (gitignored) or environment variables
> - - **Use read-only DB roles** — Create a dedicated `excel_dashboard_ro` role with SELECT-only permissions
>   - - **Encrypt stored passwords** — The `modSecurity.bas` module provides AES-256 encryption for credentials
>     - - **Use SSL connections** — Enable `sslmode=require` in connection strings for production
>       - - **Audit logging** — Enable PostgreSQL audit logging for dashboard query tracking
>        
>         - ### Create Read-Only Dashboard Role
>         - ```sql
>           -- Create dedicated read-only role for dashboard access
>           CREATE ROLE excel_dashboard_ro WITH LOGIN PASSWORD 'secure_password';
>           GRANT CONNECT ON DATABASE your_database TO excel_dashboard_ro;
>           GRANT USAGE ON SCHEMA public TO excel_dashboard_ro;
>           GRANT SELECT ON ALL TABLES IN SCHEMA public TO excel_dashboard_ro;
>           ALTER DEFAULT PRIVILEGES IN SCHEMA public
>               GRANT SELECT ON TABLES TO excel_dashboard_ro;
>           ```
>
> ---
>
> ## 🔄 Data Refresh Options
>
> | Method | Trigger | Use Case |
> |--------|---------|----------|
> | **Manual Refresh** | Button click in Excel | Ad-hoc analysis |
> | **Workbook Open** | On workbook open event | Morning briefings |
> | **Scheduled (VBA)** | Excel Application.OnTime | Intraday updates |
> | **Windows Task Scheduler** | Python scheduler.py | Overnight batch jobs |
> | **Power Automate** | Cloud trigger | Cross-platform automation |
>
> ---
>
> ## 📦 Python Dependencies
>
> ```
> psycopg2-binary>=2.9.9
> SQLAlchemy>=2.0.0
> pandas>=2.0.0
> openpyxl>=3.1.0
> xlwings>=0.30.0
> python-dotenv>=1.0.0
> schedule>=1.2.0
> cryptography>=41.0.0
> pydantic>=2.0.0
> loguru>=0.7.0
> ```
>
> ---
>
> ## 🤝 Contributing
>
> Contributions are welcome! Please follow these steps:
>
> 1. Fork the repository
> 2. 2. Create a feature branch (`git checkout -b feature/your-feature`)
>    3. 3. Commit your changes (`git commit -m 'Add: your feature description'`)
>       4. 4. Push to the branch (`git push origin feature/your-feature`)
>          5. 5. Open a Pull Request
>            
>             6. Please read `docs/CONTRIBUTING.md` for detailed guidelines.
>            
>             7. ---
>            
>             8. ## 📜 License
>
> This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.
>
> ---
>
> ## 👤 Author
>
> **Shiv Iyer**
> - GitHub: [@shiviyer](https://github.com/shiviyer)
> - - Twitter/X: [@thewebscaledba](https://twitter.com/thewebscaledba)
>   - - LinkedIn: [in/thewebscaledba](https://linkedin.com/in/thewebscaledba)
>     - - Company: [MinervaDB](https://minervadb.com) | [ChistaDATA](https://chistadata.com)
>      
>       - ---
>
> ## ⭐ Support
>
> If you find this repository helpful, please give it a ⭐ star — it motivates continued development!
>
> For enterprise support, consulting, or custom dashboard development, contact: [MinervaDB](https://minervadb.com)
>
> ---
>
> *Built with ❤️ for the PostgreSQL and Excel community*
