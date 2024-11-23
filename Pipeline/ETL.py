import pyodbc

Customer_Database = 'Customer_Database'
Customer_Table = 'Customer'
CustomerAdress_Table = 'CustomerAdress'
Purchase_Table = 'Purchase'

Connection_string = (
    'Driver={ODBC Driver 18 for SQL Server};'
    'Server=127.0.0.1,1433;'
    'UID=Admin;'
    'PWD=admin;'
    'Encrypt=no;'
)

Connect = pyodbc.connect(Connection_string, autocommit=True)
Cursor = Connect.cursor()


def Create_Databse ():

    try:
        Check_If_Database_Exist = f"SELECT COUNT(*) FROM sys.databases WHERE NAME = '{Customer_Database}'"
        Cursor.execute(Check_If_Database_Exist)
        Database_Exist =  Cursor.fetchone()[0]

        if Database_Exist:
            print(f"Database '{Customer_Database}' already exists")

        else:    
            Create_Databse_Query = f"CREATE DATABASE {Customer_Database}";
            Cursor.execute(Create_Databse_Query)
            print(f"Databasen '{Customer_Database}' was created.")
   
    except Exception as e:
        print(f"Error occurred while creating '{Customer_Database}': {e}")

def Create_Table_CustomerAdress():

    Use_Datbase = f"USE {Customer_Database}"
    Cursor.execute(Use_Datbase)

    try:
        Check_If_Table_CustomerAdress_Exist = f"SELECT Count(*) FROM sys.tables WHERE NAME = '{CustomerAdress_Table}'"
        Cursor.execute(Check_If_Table_CustomerAdress_Exist)
        Table_CustomerAdress_Exist = Cursor.fetchone()[0]

        if Table_CustomerAdress_Exist:
            print(f"Table '{CustomerAdress_Table}' already exists")

        else:
            Create_Table_CustomerAdress_Query = f"""CREATE TABLE {CustomerAdress_Table}(
            CustomerAdressID INT IDENTITY(1,1) PRIMARY KEY,
            StreetName NVARCHAR(50) NOT NULL,
            City NVARCHAR(50) NOT NULL,
            PostalCode INT NOT NULL
            )""";

            Cursor.execute(Create_Table_CustomerAdress_Query)
            print(f"Table '{CustomerAdress_Table}' was created")

    except Exception as e:
        print(f"Error occurred while creating '{CustomerAdress_Table}': {e}")

def Create_Table_Customer():

    Use_Datbase = f"USE {Customer_Database}"
    Cursor.execute(Use_Datbase)

    try:
        Check_If_Table_Customer_Exist = f"SELECT COUNT(*) FROM sys.tables WHERE NAME = '{Customer_Table}'"
        Cursor.execute(Check_If_Table_Customer_Exist)
        Table_Customer_Exist = Cursor.fetchone()[0]

        if Table_Customer_Exist:
            print(f"Table '{Customer_Table}' already exists")

        else:
            Create_Table_Customer_Query = f"""CREATE TABLE {Customer_Table}(
                CustomerID INT IDENTITY(1,1) PRIMARY KEY,
                CustomerAdressID INT NOT NULL,
                FirsnName NVARCHAR(20) NOT NULL,
                LastName NVARCHAR(20) NOT NULL,
                Phonenumber NVARCHAR(12) NOT NULL,
                Email NVARCHAR(50) NOT NULL,
                DateOfBirth DATE NOT NULL,
                StartOfMembership DATETIME NOT NULL
                )""";

            Cursor.execute(Create_Table_Customer_Query)
            print(f"Table '{Customer_Table}' was created")

            Create_Foregin_Key_CustomerAdressID = f"""ALTER TABLE {Customer_Table}
            ADD CONSTRAINT FK_Customer_CustomerAdress FOREIGN KEY (CustomerAdressID) 
            REFERENCES {CustomerAdress_Table}(CustomerAdressID)""";

            Cursor.execute(Create_Foregin_Key_CustomerAdressID)
            print(f"Foreign key for '{Customer_Table}' was added") 
    
    except Exception as e:
        print(f"Error occurred while creating '{Customer_Table}': {e}")

def Create_Table_Purchase():
    Use_Datbase = f"USE {Customer_Database}"
    Cursor.execute(Use_Datbase)

    try:
        Check_If_Table_Purchase_Exist = f"SELECT Count(*) FROM sys.tables WHERE NAME = '{Purchase_Table}'"
        Cursor.execute(Check_If_Table_Purchase_Exist)
        Table_Purchase_Exist = Cursor.fetchone()[0]

        if Table_Purchase_Exist:
            print(f"Table '{Purchase_Table}' already exists")

        else:
            Create_Table_Purchase_Query = f"""CREATE TABLE {Purchase_Table}(
            PurchaseID INT IDENTITY(1,1) PRIMARY KEY,
            CustomerID INT NOT NULL,
            Product NVARCHAR(50) NOT NULL,
            Quantity INT NOT NULL,
            PricePerProduct MONEY NOT NULL,
            TotalPrice MONEY NOT NULL,
            TimeOfPurchase DATETIME NOT NULL
            )""";

            Cursor.execute(Create_Table_Purchase_Query)
            print(f"Table '{Purchase_Table}' was created") 

            Create_Foregin_Key_CustomerID = f"""ALTER TABLE {Purchase_Table}
            ADD CONSTRAINT FK_Purchase_Customer FOREIGN KEY (CustomerID) 
            REFERENCES {Customer_Table}(CustomerID)""";

            Cursor.execute(Create_Foregin_Key_CustomerID)
            print(f"Foreign key for '{Purchase_Table}' was added") 
      
    except Exception as e:
        print(f"Error occurred while creating '{Purchase_Table}': {e}")

Create_Databse()
Create_Table_CustomerAdress()
Create_Table_Customer()
Create_Table_Purchase()

Cursor.close()
Connect.close()

 

