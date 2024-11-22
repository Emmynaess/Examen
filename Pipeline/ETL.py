import pyodbc

Connection_string = (
    'Driver={ODBC Driver 18 for SQL Server};'
    'Server=127.0.0.1,1433;'
    'UID=Admin;'
    'PWD=admin;'
    'Encrypt=no;'
)

Connect = pyodbc.connect(Connection_string, autocommit=True)
Cursor = Connect.cursor()

Customer_Database = 'Customer_Database'

def Create_Databse ():
    Check_If_Database_Exist = f"SELECT COUNT(*) FROM sys.databases WHERE NAME = '{Customer_Database}'"
    Cursor.execute(Check_If_Database_Exist)
    Database_Exist =  Cursor.fetchone()[0]

    if Database_Exist:
        print("Database already exists")

    else:    
        Create_Databse_Query = f"CREATE DATABASE {Customer_Database}";
        Cursor.execute(Create_Databse_Query)
        print(f"Databasen '{Customer_Database}' skapades.")
    
    Cursor.close()



Create_Databse()
Connect.close()



