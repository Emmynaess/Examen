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

    Create_Databse_Query = f"""IF NOT EXISTS (SELECT * FROM sys.databases WHERE NAME = '{Customer_Database}')
    BEGIN
        CREATE DATABASE {Customer_Database};
    END
    """
    
    Cursor.execute(Create_Databse_Query)
    print(f"Databasen '{Customer_Database}' skapades eller fanns redan.")
    Cursor.close()



Create_Databse()
Connect.close()

# import pyodbc

# # Inaktivera SSL-kryptering
# Connection_string = (
#     'Driver={ODBC Driver 18 for SQL Server};'
#     'Server=127.0.0.1,1433;'  # Använd port 1433 om din server är konfigurerad för den
#     'UID=Admin;'  # Ange ditt användarnamn för SQL Server Authentication
#     'PWD=admin;'  # Ange lösenordet för användaren
#     'Encrypt=no;'  # Stäng av SSL-kryptering om det orsakar problem
# )


# # Försök att skapa anslutningen
# try:
#     conn = pyodbc.connect(Connection_string)
#     print("Ansluten till SQL Server!")
    
#     # Skapa en cursor för att köra SQL-frågor
#     cursor = conn.cursor()
    
#     # Exempel på att hämta versionen på SQL Server
#     cursor.execute('SELECT @@VERSION;')
#     row = cursor.fetchone()
#     print(row)
    
#     # Stäng anslutningen och cursor
#     cursor.close()
#     conn.close()

# except pyodbc.Error as ex:
#     print(f"Fel vid anslutning: {ex}")


