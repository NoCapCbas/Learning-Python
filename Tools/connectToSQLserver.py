# sql server import
import pyodbc
# pandas import
import pandas as pd


def connectDB3():
    # connects to DB3 grabbing TDM data availability
    conn = pyodbc.connect(
                            'Driver={SQL Server};'
                            'Server=SEVENFARMS_DB3;'
                            'Database=Control;'
                            'UID=sa;'
                            'PWD=Harpua88;'
                            'Trusted_Connection=No;'
    )

    dfTDM = pd.read_sql_query(f"""
    SELECT DISTINCT
    c.[CTY_DESC] AS [CTY_RPT],
    b.[DA_ISO_CODE2] AS [CTY_ISO],
    MIN(a.[StartYM]) AS [StartYM],
    MAX(a.[StopYM]) AS [StopYM],
    a.[DA_ISO_CODE3_NUMERIC] AS [CTY_CODE]
    FROM [Control].[dbo].[Data_Availability_Monthly] a
    LEFT JOIN [Control].[dbo].[Data_Availability_Monthly] b
    ON a.[DA_ISO_CODE3_NUMERIC] = b.[DA_ISO_CODE3_NUMERIC] AND a.[StartYM] = b.[StartYM] AND a.[StopYM] = b.[StopYM]
    LEFT JOIN [SEVENFARMS_DB1].[SP_MASTER].[dbo].[CTY_MASTER] c
    ON b.[DA_ISO_CODE2] = c.[CTY_ISO]
    WHERE a.[DA_ISO_CODE3_NUMERIC] != '' AND a.[DA_ISO_CODE3_NUMERIC] IS NOT NULL
    GROUP BY a.[DA_ISO_CODE3_NUMERIC], b.[DA_ISO_CODE2], c.[CTY_DESC]""", conn)
    # removes rows with no country code
    dfTDM = dfTDM[dfTDM.CTY_CODE.notnull()]
    return dfTDM

print(connectDB3())
