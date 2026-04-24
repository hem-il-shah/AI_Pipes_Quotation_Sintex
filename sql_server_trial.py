import streamlit as st
import pandas as pd
import pyodbc

# Configuration
SERVER = 'sintex-quote-sql-dev.database.windows.net'
DATABASE = 'sintex-quote-dev-db' # Replace with your DB name
USERNAME = 'swa-pipes-admin-login'      # Replace with your ID
PASSWORD = 'SintexApps@123'      # Replace with your Password
DRIVER = '{ODBC Driver 17 for SQL Server}'

def get_connection():
    conn_str = (
        f'DRIVER={DRIVER};'
        f'SERVER={SERVER};'
        f'PORT=1433;'
        f'DATABASE={DATABASE};'
        f'UID={USERNAME};'
        f'PWD={PASSWORD};'
        f'Encrypt=yes;'
        f'TrustServerCertificate=no;'
        f'Connection Timeout=30;'
    )
    return pyodbc.connect(conn_str)

def main():
    st.set_page_config(page_title="SQL Data Viewer", layout="wide")
    st.title("📊 SQL Server Data Explorer")

    query = st.text_area("Enter your SQL Query:", "SELECT TOP 10 * FROM sys.tables")

    if st.button("Run Query"):
        try:
            with get_connection() as conn:
                df = pd.read_sql(query, conn)
                st.success(f"Returned {len(df)} rows")
                st.dataframe(df, use_container_width=True)
        except Exception as e:
            st.error(f"Error: {e}")

if __name__ == "__main__":
    main()