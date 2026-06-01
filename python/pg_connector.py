#!/usr/bin/env python3
"""
pg_connector.py
Core PostgreSQL Connector for Excel-PostgreSQL-Dashboard
Author: Shiv Iyer | MinervaDB | ChistaDATA
License: MIT
"""

import os
import configparser
from typing import Optional, Union, Dict, Any, List
from contextlib import contextmanager
from pathlib import Path

import psycopg2
import psycopg2.extras
from psycopg2 import pool
import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.exc import SQLAlchemyError
from dotenv import load_dotenv
from loguru import logger

load_dotenv()


class PostgreSQLConfig:
      """Configuration container for PostgreSQL connection parameters."""

    def __init__(self, host="localhost", port=5432, database="",
                                  username="", password="", schema="public",
                                  sslmode="prefer", connect_timeout=30,
                                  pool_min=1, pool_max=10,
                                  application_name="Excel-PostgreSQL-Dashboard"):
                                            self.host = host
                                            self.port = port
                                            self.database = database
                                            self.username = username
                                            self.password = password
                                            self.schema = schema
                                            self.sslmode = sslmode
                                            self.connect_timeout = connect_timeout
                                            self.pool_min = pool_min
                                            self.pool_max = pool_max
                                            self.application_name = application_name

    @classmethod
    def from_config_file(cls, config_path):
              config = configparser.ConfigParser()
              config_path = Path(config_path)
              if not config_path.exists():
                            raise FileNotFoundError(f"Config file not found: {config_path}")
                        config.read(config_path)
        pg = config["postgresql"]
        return cls(
                      host=pg.get("host", "localhost"),
                      port=pg.getint("port", 5432),
                      database=pg.get("database", ""),
                      username=pg.get("username", ""),
                      password=pg.get("password", ""),
                      schema=pg.get("schema", "public"),
                      sslmode=pg.get("sslmode", "prefer"),
                      connect_timeout=pg.getint("connect_timeout", 30),
        )

    @classmethod
    def from_env(cls):
              return cls(
                  host=os.getenv("PG_HOST", "localhost"),
                  port=int(os.getenv("PG_PORT", "5432")),
                  database=os.getenv("PG_DATABASE", ""),
                  username=os.getenv("PG_USERNAME", ""),
                  password=os.getenv("PG_PASSWORD", ""),
                  schema=os.getenv("PG_SCHEMA", "public"),
                  sslmode=os.getenv("PG_SSLMODE", "prefer"),
    )

    @property
    def dsn(self):
              return (
                  f"host={self.host} port={self.port} dbname={self.database} "
                  f"user={self.username} password={self.password} "
                  f"sslmode={self.sslmode} connect_timeout={self.connect_timeout} "
                  f"application_name={self.application_name} "
                  f"options=-c search_path={self.schema}"
    )

    @property
    def sqlalchemy_url(self):
              return (
                  f"postgresql+psycopg2://{self.username}:{self.password}"
                  f"@{self.host}:{self.port}/{self.database}"
    )


class PostgreSQLConnector:
      """
          Enterprise-grade PostgreSQL connector with connection pooling.

              Usage:
                      conn = PostgreSQLConnector.from_config("config/config.ini")
                              df = conn.query_to_dataframe("SELECT * FROM vw_executive_kpis")
                                      conn.close()
                                          """

    def __init__(self, config: PostgreSQLConfig):
              self.config = config
        self._pool = None
        self._engine = None
        self._is_connected = False

    @classmethod
    def from_config(cls, config_path):
              config = PostgreSQLConfig.from_config_file(config_path)
        connector = cls(config)
        connector.connect()
        return connector

    @classmethod
    def from_env(cls):
              config = PostgreSQLConfig.from_env()
        connector = cls(config)
        connector.connect()
        return connector

    def connect(self):
              try:
                            self._pool = pool.ThreadedConnectionPool(
                                              minconn=self.config.pool_min,
                                              maxconn=self.config.pool_max,
                                              dsn=self.config.dsn,
                            )
                            self._engine = create_engine(
                                self.config.sqlalchemy_url,
                                pool_pre_ping=True,
                                connect_args={
                                    "options": f"-c search_path={self.config.schema}",
                                    "sslmode": self.config.sslmode,
                                },
                            )
                            self._is_connected = True
                            logger.success(f"Connected: {self.config.database}@{self.config.host}")
                            return True
except Exception as e:
            logger.error(f"Connection failed: {e}")
            self._is_connected = False
            return False

    def disconnect(self):
              if self._pool:
                            self._pool.closeall()
                        if self._engine:
                                      self._engine.dispose()
                                  self._is_connected = False

    def close(self):
              self.disconnect()

    @contextmanager
    def get_connection(self):
              conn = self._pool.getconn()
        try:
                      yield conn
                      conn.commit()
except Exception:
            conn.rollback()
            raise
finally:
            self._pool.putconn(conn)

    def query_to_dataframe(self, sql, params=None):
              """Execute SELECT query and return Pandas DataFrame."""
        with self._engine.connect() as conn:
                      return pd.read_sql(text(sql), conn, params=params)

    def query_scalar(self, sql, params=None, default=None):
              """Return first column of first row."""
        try:
                      with self.get_connection() as conn:
                                        with conn.cursor() as cur:
                                                              cur.execute(sql, params)
                                                              row = cur.fetchone()
                                                              return row[0] if row else default
        except Exception as e:
            logger.error(f"Scalar query failed: {e}")
            return default

    def query_fetchall(self, sql, params=None, as_dict=True):
              """Fetch all rows as list of dicts or tuples."""
        with self.get_connection() as conn:
                      factory = psycopg2.extras.RealDictCursor if as_dict else None
                      with conn.cursor(cursor_factory=factory) as cur:
                                        cur.execute(sql, params)
                                        return cur.fetchall()

              def execute_non_query(self, sql, params=None):
                        """Execute INSERT/UPDATE/DELETE. Returns row count."""
                        with self.get_connection() as conn:
                                      with conn.cursor() as cur:
                                                        cur.execute(sql, params)
                                                        return cur.rowcount

                              def execute_many(self, sql, data):
                                        """Batch INSERT/UPDATE using execute_batch."""
                                        with self.get_connection() as conn:
                                                      with conn.cursor() as cur:
                                                                        psycopg2.extras.execute_batch(cur, sql, data, page_size=1000)
                                                                        return len(data)

                                              def query_from_file(self, sql_file_path, params=None):
                                                        """Load SQL from file and return DataFrame."""
                                                        sql_path = Path(sql_file_path)
                                                        if not sql_path.exists():
                                                                      raise FileNotFoundError(f"SQL file not found: {sql_path}")
                                                                  return self.query_to_dataframe(sql_path.read_text(encoding="utf-8"), params)

    def get_kpi_snapshot(self, kpi_queries: Dict[str, str]):
              """Execute multiple KPI scalar queries. Returns {name: value} dict."""
        results = {}
        for name, sql in kpi_queries.items():
                      try:
                                        results[name] = self.query_scalar(sql)
except Exception as e:
                  logger.warning(f"KPI '{name}' failed: {e}")
                  results[name] = None
          return results

    def get_dashboard_data(self, queries: Dict[str, str]):
              """Execute multiple named queries. Returns {name: DataFrame} dict."""
              datasets = {}
              for name, sql in queries.items():
                            try:
                                              datasets[name] = self.query_to_dataframe(sql)
                                              logger.info(f"Dataset '{name}': {len(datasets[name])} rows")
except Exception as e:
                logger.error(f"Dataset '{name}' failed: {e}")
                datasets[name] = pd.DataFrame()
        return datasets

    def list_schemas(self):
              df = self.query_to_dataframe(
                            "SELECT schema_name FROM information_schema.schemata "
                            "WHERE schema_name NOT IN ('pg_catalog','information_schema') "
                            "ORDER BY schema_name"
              )
              return df["schema_name"].tolist()

    def list_tables(self, schema="public"):
              return self.query_to_dataframe(
                            "SELECT table_name, table_type FROM information_schema.tables "
                            "WHERE table_schema = %(schema)s ORDER BY table_name",
                            params={"schema": schema},
              )

    def get_server_version(self):
              return self.query_scalar("SELECT version()") or "Unknown"

    def __enter__(self):
              return self

    def __exit__(self, exc_type, exc_val, exc_tb):
              self.disconnect()
              return False

    def __repr__(self):
              status = "connected" if self._is_connected else "disconnected"
              return f"PostgreSQLConnector({self.config.database}@{self.config.host}, {status})"


def connect(config_path=None):
      """Quick-connect factory. Uses config file or env variables."""
      if config_path:
                return PostgreSQLConnector.from_config(config_path)
            return PostgreSQLConnector.from_env()


if __name__ == "__main__":
      import sys
    cfg = sys.argv[1] if len(sys.argv) > 1 else "config/config.ini"
    with connect(cfg) as conn:
              print(f"PostgreSQL Version: {conn.get_server_version()}")
              print(f"Schemas: {conn.list_schemas()}")
