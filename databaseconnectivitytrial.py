import streamlit as st
import pyodbc
import pandas as pd

# Set up page config
st.set_page_config(page_title="SQL Server Query Interface", layout="wide")

# Connection parameters based on your screenshots
SERVER = 'HEMIL\\SQLEXPRESS'
DATABASE = 'sintex_quote_db'

def get_connection():
    """Establishes a connection using Windows Authentication."""
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        f"Trusted_Connection=yes;"
    )
    return pyodbc.connect(conn_str)

st.title("🛢️ SQL Server Query Interface")
st.info(f"Connected to: **{SERVER}** | Database: **{DATABASE}**")

# SQL Input Area
query = st.text_area("Enter your SQL query here:", height=200, placeholder="SELECT * FROM dbo.ocr_logs")

if st.button("Run Query"):
    if query.strip():
        try:
            conn = get_connection()
            
            # Use pandas to read the SQL query directly into a dataframe
            df = pd.read_sql(query, conn)
            
            st.success("Query executed successfully!")
            st.subheader("Results")
            st.dataframe(df, use_container_width=True)
            
            # Option to download results
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button("Download as CSV", csv, "query_results.csv", "text/csv")
            
            conn.close()
        except Exception as e:
            st.error(f"Error: {e}")
    else:
        st.warning("Please enter a query first.")

# Sidebar - Database Schema Reference (Optional)
with st.sidebar:
    st.header("Schema Reference")
    st.write("**Tables in sintex_quote_db:**")
    st.code("dbo.ocr_logs\ndbo.quotation_logs")