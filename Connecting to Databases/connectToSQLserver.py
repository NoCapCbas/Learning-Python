# pip install pyodbc
# sql server import
import pyodbc


def connectDB3():
    
    conn = pyodbc.connect(
                            'Driver={SQL Server};'  # or 'Driver={ODBC Driver 17 for SQL Server};'
                            'Server=server_name;'
                            'Database=database_name;'
                            'UID=username;'
                            'PWD=password;'
                            'Trusted_Connection=No;'
    )

    df = pd.read_sql_query(f"""
    SELECT *
    FROM TABLE_NAME;
    """, conn)

    return df

print(connectDB3())
