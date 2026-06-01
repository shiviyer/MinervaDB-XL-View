# 🪟 MinervaDB XL View — Windows Installation Guide

Complete guide for installing and running MinervaDB XL View on Windows 10/11 with full Excel VBA + PostgreSQL ODBC connectivity.

---

## 📋 Windows Prerequisites

| Component | Minimum | Recommended | Notes |
|-----------|---------|-------------|-------|
| Windows | 10 (64-bit) | 11 (64-bit) | 32-bit not supported |
| Microsoft Excel | 2016 | Microsoft 365 | 64-bit version recommended |
| PostgreSQL ODBC | 13.x | 16.x | psqlODBC 64-bit |
| Python | 3.9 | 3.12 | 64-bit installer |
| Git | Any | Latest | git-scm.com |
| .NET Framework | 4.7.2 | 4.8 | Usually pre-installed |

---

## 🚀 Quick Install (Automated PowerShell)

Open **PowerShell as Administrator** and run:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
git clone https://github.com/shiviyer/MinervaDB-XL-View.git
cd MinervaDB-XL-View
.\install\install_windows.ps1
```

---

## 🔧 Manual Installation Steps

### Step 1: Install Git for Windows

Download and install from https://git-scm.com/download/win

Or using winget:
```powershell
winget install --id Git.Git -e --source winget
```

---

### Step 2: Install Python 3.12 (64-bit)

Download the **64-bit installer** from https://www.python.org/downloads/windows/

Important installer options:
- ✅ Check **"Add Python to PATH"**
- ✅ Check **"Install for all users"** (recommended)
- Click **"Customize installation"** → check pip, tcl/tk, py launcher

Or using winget:
```powershell
winget install --id Python.Python.3.12 -e --source winget
```

Verify:
```powershell
python --version    # Should show Python 3.12.x
pip --version
```

---

### Step 3: Install PostgreSQL ODBC Driver (psqlODBC)

1. Go to https://www.postgresql.org/ftp/odbc/versions/msi/
2. Download the latest **psqlodbc_xx_xx_xxxx-x64.zip** (64-bit)
3. Extract and run the **.msi** installer
4. Accept defaults and complete installation

Or using Chocolatey (if installed):
```powershell
choco install psqlodbc
```

Verify installation:
1. Press **Win+R**, type `odbcad32`, press Enter
2. Go to **Drivers** tab
3. Confirm "PostgreSQL ANSI" and "PostgreSQL Unicode" are listed

---

### Step 4: Clone MinervaDB XL View

Open PowerShell:
```powershell
cd C:\Users\YourName\Documents
git clone https://github.com/shiviyer/MinervaDB-XL-View.git
cd MinervaDB-XL-View
```

---

### Step 5: Set Up Python Virtual Environment

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install -r python\requirements.txt
```

If you get a script execution error:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Verify packages:
```powershell
pip list | Select-String "psycopg2|pandas|openpyxl|sqlalchemy"
```

---

### Step 6: Configure Database Connection

```powershell
Copy-Item config\config.example.ini config\config.ini
notepad config\config.ini
```

Update the configuration:
```ini
[database]
host     = your-postgresql-host
port     = 5432
database = your_database_name
username = excel_dashboard_ro
password = your_secure_password
ssl_mode = require

[odbc]
dsn_name = MinervaDB_XL_View
driver   = PostgreSQL Unicode

[output]
output_dir = .\output
```

> **Security:** Never commit config\config.ini to git. It is already in .gitignore.

---

### Step 7: Create ODBC Data Source (DSN)

This step is required for the Excel VBA connection.

**Method A: Using the ODBC Administrator (GUI)**
1. Press **Win+R**, type `odbcad32` (for 64-bit) or `C:\Windows\SysWOW64\odbcad32.exe` (for 32-bit Excel)
2. Click the **User DSN** tab → click **Add**
3. Select **PostgreSQL Unicode** → click **Finish**
4. Fill in:
   - **Data Source**: MinervaDB_XL_View
   - **Description**: MinervaDB XL View PostgreSQL Connection
   - **Server**: your-postgresql-host
   - **Port**: 5432
   - **Database**: your_database_name
   - **User Name**: excel_dashboard_ro
5. Click **Test** to verify, then **Save**

**Method B: Using PowerShell (Automated)**
```powershell
# Run the install script which creates the DSN automatically
.\install\install_windows.ps1
```

---

### Step 8: Set Up Sample Database (Optional)

If you have psql available:
```powershell
# Set environment variable for password
$env:PGPASSWORD = "your_postgres_admin_password"

# Create schema
psql -h your-host -U postgres -d your_database -f sql\setup\create_sample_schema.sql

# Seed data
psql -h your-host -U postgres -d your_database -f sql\setup\seed_sample_data.sql
```

---

### Step 9: Test Python Connection

```powershell
.\.venv\Scripts\Activate.ps1
python python\pg_connector.py
```

Expected output:
```
[INFO] MinervaDB XL View: Testing connection...
[INFO] MinervaDB XL View: Connected to PostgreSQL 16.x
[INFO] MinervaDB XL View: Connection test passed!
```

---

### Step 10: Import VBA Modules into Excel

1. Open Microsoft Excel
2. Press **Alt+F11** to open the VBA Editor
3. Go to **File** → **Import File...**
4. Navigate to the `vba\` folder
5. Import these files in order:
   - `modErrorHandler.bas`
   - `modConnection.bas`
   - `modQueryRunner.bas`
   - `modDataRefresh.bas`
   - `modFormatting.bas`
   - `modSecurity.bas`
   - `modExport.bas`
6. Go to **Tools** → **References**
7. Check **"Microsoft ActiveX Data Objects 6.1 Library"**
8. Click **OK**

---

### Step 11: Set Up Excel Connection Configuration

In the Excel VBA Immediate Window (**Ctrl+G** in VBA Editor):
```vba
' Initialize the connection with your DSN
Call modConnection.SetConnectionString("DSN=MinervaDB_XL_View;UID=excel_dashboard_ro;PWD=yourpassword")

' Test connection
Call modConnection.TestConnection()
```

---

### Step 12: Test Full Dashboard Refresh

In the VBA Immediate Window:
```vba
' Initialize the refresh module
Call modDataRefresh.InitializeRefresh(15)

' Refresh all dashboards
Call modDataRefresh.RefreshAllDashboards()
```

---

## ⏰ Set Up Scheduled Auto-Refresh (Windows Task Scheduler)

**Method A: PowerShell Script**
```powershell
# Create a scheduled task to refresh every 15 minutes
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument @"
-NonInteractive -WindowStyle Hidden -Command "
  cd 'C:\Users\YourName\Documents\MinervaDB-XL-View'
  .\.venv\Scripts\Activate.ps1
  python python\etl_pipeline.py --pipeline all
"
"@

$trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 15) -Once -At (Get-Date)
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

Register-ScheduledTask -TaskName "MinervaDB XL View Refresh" `
  -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest
```

**Method B: Task Scheduler GUI**
1. Press **Win+R**, type `taskschd.msc`
2. Click **Create Basic Task**
3. Name: "MinervaDB XL View Refresh"
4. Trigger: Daily, repeat every 15 minutes
5. Action: Start a program
6. Program: `powershell.exe`
7. Arguments: `-File "C:\path\to\MinervaDB-XL-View\install\refresh.ps1"`

---

## 🔑 Environment Variables (Security Best Practice)

Set system environment variables instead of storing credentials in config.ini:

```powershell
# Set environment variables (run as Administrator)
[System.Environment]::SetEnvironmentVariable("MINERVADB_HOST", "your-host", "User")
[System.Environment]::SetEnvironmentVariable("MINERVADB_PORT", "5432", "User")
[System.Environment]::SetEnvironmentVariable("MINERVADB_DB", "your_database", "User")
[System.Environment]::SetEnvironmentVariable("MINERVADB_USER", "excel_dashboard_ro", "User")
[System.Environment]::SetEnvironmentVariable("MINERVADB_PASSWORD", "your_password", "User")
```

---

## 🧪 Run Tests

```powershell
.\.venv\Scripts\Activate.ps1
pip install pytest
pytest tests\ -v
```

---

## ❓ Common Windows Issues

| Issue | Solution |
|-------|----------|
| "Data source name not found" | Run odbcad32 and create the MinervaDB_XL_View DSN |
| "ActiveX component can't create object" | Tools → References → check ADO 6.1 Library |
| psycopg2 install fails | Use `pip install psycopg2-binary` |
| PowerShell script blocked | Run `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` |
| 32-bit vs 64-bit ODBC mismatch | Use SysWOW64\odbcad32.exe if Excel is 32-bit |
| "Login failed for user" | Check pg_hba.conf on PostgreSQL server |
| SSL connection error | Set ssl_mode = disable or configure server SSL |

See [TROUBLESHOOTING.md](TROUBLESHOOTING.md) for more detailed solutions.

---

*MinervaDB XL View — Enterprise PostgreSQL Analytics Platform*
*Windows Installation Guide v1.0*
