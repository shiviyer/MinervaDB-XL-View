"""
MinervaDB XL View - ETL Pipeline Module
etl_pipeline.py
============================================================
Enterprise-grade ETL pipeline for extracting data from
PostgreSQL and loading it into Excel-ready formats.
============================================================
"""

import logging
import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
from typing import Optional, Dict, List, Any, Tuple
from pathlib import Path
import json

from pg_connector import MinervaDBConnector

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] MinervaDB XL View ETL: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger('minervadb_xl_view.etl')


class MinervaDBETLPipeline:
    """
    MinervaDB XL View ETL Pipeline
    
    Orchestrates data extraction from PostgreSQL, transformation,
    validation, and loading into Excel dashboard formats.
    """

    def __init__(self, config_path: str = 'config/config.ini'):
        self.connector = MinervaDBConnector(config_path)
        self.run_id = datetime.now().strftime('%Y%m%d_%H%M%S')
        self.metrics: Dict[str, Any] = {
            'run_id': self.run_id,
            'start_time': datetime.now().isoformat(),
            'pipelines_run': 0,
            'rows_processed': 0,
            'errors': []
        }
        logger.info(f"MinervaDB XL View ETL Pipeline initialized. Run ID: {self.run_id}")

    # =========================================================
    # Sales Dashboard ETL
    # =========================================================
    def run_sales_pipeline(self, output_path: Optional[str] = None) -> pd.DataFrame:
        """Extract, transform, and load sales KPI data."""
        logger.info("Starting Sales Dashboard ETL pipeline...")

        with self.connector.get_connection() as conn:
            df_raw = self.connector.execute_query(conn, """
                SELECT
                    s.sale_date,
                    s.region,
                    s.product_category,
                    s.salesperson,
                    s.quantity,
                    s.unit_price,
                    s.discount_pct,
                    s.revenue,
                    s.cost,
                    s.gross_profit,
                    c.customer_name,
                    c.customer_segment
                FROM sales_transactions s
                JOIN customers c ON s.customer_id = c.customer_id
                WHERE s.sale_date >= CURRENT_DATE - INTERVAL '90 days'
                ORDER BY s.sale_date DESC
            """)

            if df_raw is None or df_raw.empty:
                logger.warning("No sales data returned.")
                return pd.DataFrame()

            df = self._transform_sales(df_raw)

            if output_path:
                self._load_to_excel(df, output_path, 'Sales_Data')
                self._load_sales_summary(df, output_path)

            self.metrics['pipelines_run'] += 1
            self.metrics['rows_processed'] += len(df)
            logger.info(f"Sales ETL complete. {len(df)} rows processed.")
            return df

    def _transform_sales(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply sales-specific transformations."""
        df = df.copy()
        df['sale_date'] = pd.to_datetime(df['sale_date'])
        df['year'] = df['sale_date'].dt.year
        df['month'] = df['sale_date'].dt.month
        df['quarter'] = df['sale_date'].dt.quarter
        df['revenue'] = pd.to_numeric(df['revenue'], errors='coerce').fillna(0)
        df['cost'] = pd.to_numeric(df['cost'], errors='coerce').fillna(0)
        df['gross_profit'] = df['revenue'] - df['cost']
        df['gross_margin_pct'] = np.where(
            df['revenue'] > 0,
            (df['gross_profit'] / df['revenue'] * 100).round(2),
            0
        )
        return df

    def _load_sales_summary(self, df: pd.DataFrame, output_path: str):
        """Create a summary pivot for the Excel dashboard."""
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            monthly = df.groupby(['year', 'month']).agg(
                revenue=('revenue', 'sum'),
                transactions=('revenue', 'count'),
                gross_profit=('gross_profit', 'sum')
            ).reset_index()
            monthly['gross_margin_pct'] = (monthly['gross_profit'] / monthly['revenue'] * 100).round(2)
            monthly.to_excel(writer, sheet_name='Sales_Monthly', index=False)

    # =========================================================
    # Finance Dashboard ETL
    # =========================================================
    def run_finance_pipeline(self, output_path: Optional[str] = None) -> pd.DataFrame:
        """Extract, transform, and load financial KPI data."""
        logger.info("Starting Finance Dashboard ETL pipeline...")

        with self.connector.get_connection() as conn:
            df_raw = self.connector.execute_query(conn, """
                SELECT
                    period_date,
                    account_category,
                    account_name,
                    actual_amount,
                    budget_amount,
                    prior_year_amount,
                    department
                FROM financial_data
                WHERE period_date >= DATE_TRUNC('year', CURRENT_DATE) - INTERVAL '1 year'
                ORDER BY period_date DESC
            """)

            if df_raw is None or df_raw.empty:
                return pd.DataFrame()

            df = self._transform_finance(df_raw)
            if output_path:
                self._load_to_excel(df, output_path, 'Finance_Data')
            self.metrics['pipelines_run'] += 1
            self.metrics['rows_processed'] += len(df)
            return df

    def _transform_finance(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply finance-specific transformations."""
        df = df.copy()
        df['period_date'] = pd.to_datetime(df['period_date'])
        df['year'] = df['period_date'].dt.year
        df['month'] = df['period_date'].dt.month
        for col in ['actual_amount', 'budget_amount', 'prior_year_amount']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        df['variance_vs_budget'] = df['actual_amount'] - df['budget_amount']
        df['variance_pct_budget'] = np.where(
            df['budget_amount'] != 0,
            (df['variance_vs_budget'] / df['budget_amount'].abs() * 100).round(2),
            0
        )
        return df

    # =========================================================
    # HR Dashboard ETL
    # =========================================================
    def run_hr_pipeline(self, output_path: Optional[str] = None) -> pd.DataFrame:
        """Extract, transform, and load HR KPI data."""
        logger.info("Starting HR Dashboard ETL pipeline...")

        with self.connector.get_connection() as conn:
            df_raw = self.connector.execute_query(conn, """
                SELECT
                    e.employee_id,
                    e.hire_date,
                    e.department,
                    e.job_title,
                    e.employment_status,
                    e.salary,
                    e.performance_rating,
                    e.location
                FROM employees e
                WHERE e.employment_status IN ('Active', 'Terminated')
                ORDER BY e.department, e.hire_date
            """)

            if df_raw is None or df_raw.empty:
                return pd.DataFrame()

            df = self._transform_hr(df_raw)
            if output_path:
                self._load_to_excel(df, output_path, 'HR_Data')
            self.metrics['pipelines_run'] += 1
            self.metrics['rows_processed'] += len(df)
            return df

    def _transform_hr(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply HR-specific transformations."""
        df = df.copy()
        today = pd.Timestamp.today()
        df['hire_date'] = pd.to_datetime(df['hire_date'])
        df['tenure_days'] = (today - df['hire_date']).dt.days
        df['tenure_years'] = (df['tenure_days'] / 365.25).round(2)
        df['salary'] = pd.to_numeric(df['salary'], errors='coerce').fillna(0)
        df['is_active'] = df['employment_status'] == 'Active'
        return df

    # =========================================================
    # Full Refresh
    # =========================================================
    def run_all_pipelines(self, output_dir: str = 'output') -> Dict[str, pd.DataFrame]:
        """Run all ETL pipelines."""
        logger.info("MinervaDB XL View: Starting full ETL refresh...")
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        output_path = f"{output_dir}/MinervaDB_XL_View_{self.run_id}.xlsx"
        results = {}
        pipelines = [('sales', self.run_sales_pipeline), ('finance', self.run_finance_pipeline), ('hr', self.run_hr_pipeline)]
        for name, fn in pipelines:
            try:
                results[name] = fn(output_path)
            except Exception as e:
                logger.error(f"Pipeline {name!r} failed: {e}")
                self.metrics['errors'].append({'pipeline': name, 'error': str(e)})
                results[name] = pd.DataFrame()
        self.metrics['end_time'] = datetime.now().isoformat()
        self._save_run_metrics(output_dir)
        return results

    def _load_to_excel(self, df: pd.DataFrame, output_path: str, sheet_name: str):
        """Load DataFrame to Excel sheet."""
        path = Path(output_path)
        if path.exists():
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        logger.info(f"Loaded {len(df)} rows to {sheet_name!r}")

    def _save_run_metrics(self, output_dir: str):
        """Save ETL run metrics to JSON."""
        metrics_path = f"{output_dir}/etl_metrics_{self.run_id}.json"
        with open(metrics_path, 'w') as f:
            json.dump(self.metrics, f, indent=2, default=str)


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description="MinervaDB XL View ETL Pipeline")
    parser.add_argument('--config', default='config/config.ini')
    parser.add_argument('--output', default='output')
    parser.add_argument('--pipeline', default='all', choices=['all', 'sales', 'finance', 'hr'])
    args = parser.parse_args()
    etl = MinervaDBETLPipeline(config_path=args.config)
    if args.pipeline == 'all':
        etl.run_all_pipelines(output_dir=args.output)
    elif args.pipeline == 'sales':
        etl.run_sales_pipeline()
    elif args.pipeline == 'finance':
        etl.run_finance_pipeline()
    elif args.pipeline == 'hr':
        etl.run_hr_pipeline()
