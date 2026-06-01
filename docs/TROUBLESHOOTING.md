# MinervaDB XL View - Troubleshooting Guide

This guide covers common issues and solutions when working with MinervaDB XL View.

---

## Connection Issues

### Error: "Data source name not found and no default driver specified"

**Cause:** The PostgreSQL ODBC driver is not installed or not configured.

**Solution:**
1. Download and install psqlODBC from https://www.postgresql.org/ftp/odbc/versions/
2. Open Windows ODBC Data Source Administrator (odbcad32.exe)
3. Add a new User DSN using the PostgreSQL Unicode driver
4. Configure: Server, Port (5432), Database, Username, Password
5. Click "Test" to verify the connection
6. Update config/config.ini with your DSN name

---

### Error: "Connection refused to host:port"

**Cause:** PostgreSQL server is not running or not accessible.

**Solution:**
1. Verify PostgreSQL is running: `pg_ctl status`
2. Check the host and port in config/config.ini
3. Verify firewall rules allow TCP port 5432
4. Test connectivity: `telnet your-postgres-host 5432`
5. Check PostgreSQL listen_addresses in postgresql.conf

---

### Error: "Password authentication failed for user"

**Cause:** Incorrect username or password in configuration.

**Solution:**
1. Verify credentials in config/config.ini
2. Test login directly: `psql -h host -U username -d database`
3. Check pg_hba.conf allows the connection method
4. Reset password if needed: `ALTER USER username WITH PASSWORD 'newpass';`

---

## Excel VBA Issues

### Error: "Compile error: User-defined type not defined"

**Cause:** ADODB reference is not set in the VBA project.

**Solution:**
1. Open Excel VBA Editor (Alt+F11)
2. Go to Tools > References
3. Check "Microsoft ActiveX Data Objects 6.1 Library"
4. Click OK and try again

---

### Error: "Run-time error '429': ActiveX component can't create object"

**Cause:** ADODB is not installed or registered.

**Solution:**
1. Verify MDAC (Microsoft Data Access Components) is installed
2. Re-register ADODB: `regsvr32 msado15.dll`
3. Ensure 64-bit Excel uses 64-bit ODBC drivers

---

### Dashboard shows no data after refresh

**Cause:** Queries returning no results or writing to wrong cells.

**Solution:**
1. Run the query directly in psql to verify it returns data
2. Check sheet name constants in modDataRefresh.bas match your actual sheet names
3. Check the ErrorLog sheet in your workbook for logged errors
4. Enable debug mode: add `Debug.Print` statements in VBA

---

## Python Issues

### ImportError: No module named 'psycopg2'

```bash
pip install psycopg2-binary
```

### ImportError: No module named 'openpyxl'

```bash
pip install -r python/requirements.txt
```

### ConnectionError: could not connect to server

1. Verify config/config.ini has correct connection details
2. Test: `python -c "import psycopg2; psycopg2.connect(host='host', dbname='db', user='user', password='pass')"`
3. Check network and firewall settings

---

## SQL Issues

### ERROR: relation does not exist

```bash
psql -U postgres -d your_database -f sql/setup/create_sample_schema.sql
psql -U postgres -d your_database -f sql/setup/seed_sample_data.sql
```

### ERROR: permission denied for table

```sql
GRANT SELECT ON ALL TABLES IN SCHEMA public TO excel_dashboard_ro;
GRANT SELECT ON ALL SEQUENCES IN SCHEMA public TO excel_dashboard_ro;
```

### Queries running slowly

1. Run EXPLAIN ANALYZE on slow queries
2. Create indexes: `psql -f sql/setup/create_indexes.sql`
3. Run VACUUM ANALYZE on large tables

---

## Performance Tips

- Set refresh interval to at least 5 minutes
- Use the read-only role (excel_dashboard_ro) to prevent accidental writes
- Add connection pooling for multiple concurrent users
- Use parameterized date ranges to limit data volume

---

## Getting Help

1. Check the ErrorLog sheet in your Excel workbook
2. Review MinervaDB_XL_View_Errors.log in the workbook directory
3. Run Python tests: `pytest tests/`
4. Open an issue at https://github.com/shiviyer/MinervaDB-XL-View/issues

---

*MinervaDB XL View - Enterprise PostgreSQL Analytics Platform*
