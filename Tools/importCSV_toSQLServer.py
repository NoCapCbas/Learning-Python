import pandas as pd
import pyodbc

# Import CSV
data = pd.read_csv (r'C:\Users\DDiaz\Documents\Dropbox\Exports_2018.csv')
df = pd.DataFrame(data)
# print(df)

# Connect to SQL Server
conn = pyodbc.connect(
                        'Driver={SQL Server};'
                        'Server=SEVENFARMS_DB1;'
                        'Database=Andrews_practice_database;'
                        'UID=sa;'
                        'PWD=Harpua88;'
                        'Trusted_Connection=No;'
)
cursor = conn.cursor()

# Create Table
cursor.execute('''
		CREATE TABLE products (
			TRADEFLOWS VARCHAR(250),
			PERIOD VARCHAR(6),
            HScode VARCHAR(12)
			)
               ''')

# Insert DataFrame to Table
for row in df.itertuples():
    cursor.execute('''
                INSERT INTO products (TRADEFLOWS, PERIOD, HScode)
                VALUES (?,?,?)
                ''',
                row.TradeFlows,
                row.STAT_YM,
                row.HScode
                )
conn.commit()
