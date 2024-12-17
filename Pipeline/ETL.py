import pyodbc
import os
import pandas as pd
import re
import gc
from datetime import datetime

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

current_dir = os.getcwd()
Kunddata_Webbshop = os.path.join(current_dir, 'kunddata_webbshop.xlsx') 
Control_Data = os.path.join(current_dir, 'kunddata_adresser_kontroll.xlsx')
Slaskfil = os.path.join(current_dir, 'slaskfil.xlsx')
kontrollfil = os.path.join(current_dir, 'kontrollfil.xlsx')

def Create_Database ():

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

def Create_Slask():

    Slask_data = Webshop_Data[Webshop_Data['Slask'] == True]
    valid_data = Webshop_Data[Webshop_Data['Slask'] == False]

    if not Slask_data.empty:
        if os.path.exists(Slaskfil):
            existing_Slask = pd.read_excel(Slaskfil)
            Slask_data = pd.concat([existing_Slask, Slask_data], ignore_index=True)
        else:
            existing_Slask = pd.DataFrame()
            Slask_data = pd.concat([existing_Slask, Slask_data], ignore_index=True)

        Slask_data.to_excel(Slaskfil, index=False, engine='openpyxl')
        print(f"Slask file '{Slaskfil}' has been updated.")
    else:
        print("No new invalid rows found.")

def Control_Names(Webshop_Data):
    for index, row in Webshop_Data.iterrows():
        customer_name = row['Kundnamn']

        if pd.isnull(customer_name) or not customer_name.strip():
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Missing customer name sent to slask.")
            continue

        if not re.match(r"^[a-zA-ZåäöÅÄÖéÉíÍóÓúÚüÜñÑ' -]+$", customer_name.strip()):
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Customer name '{customer_name}' contains invalid characters. Sent to slask.")
            continue

        name_parts = customer_name.strip().split()
        if len(name_parts) < 2:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Invalid name '{customer_name}' (missing first or last name). Sent to slask.")
        else:
            Webshop_Data.at[index, 'FirstName'] = name_parts[0].capitalize()
            Webshop_Data.at[index, 'LastName'] = ' '.join(part.capitalize() for part in name_parts[1:])

def Control_Birthdate(Webshop_Data):
    for index, row in Webshop_Data.iterrows():
        birthdate = row['Födelsedatum']

        if pd.isnull(birthdate) or not str(birthdate).strip():
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Missing birthdate sent to slask.")
            continue

        try:
            birthdate = pd.to_datetime(birthdate, format='%Y-%m-%d', errors='raise')
            age = (pd.Timestamp.now() - birthdate).days // 365
            
            if age < 18:
                Webshop_Data.at[index, 'Slask'] = True
                print(f"Row {index}: Underage customer sent to slask.")
        
        except Exception:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Invalid birthdate format '{birthdate}' sent to slask.")

def Control_Email(Webshop_Data):
    for index, row in Webshop_Data.iterrows():
        email = row['Email']

        if pd.isnull(email) or not str(email).strip():
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Missing email sent to slask.")
            continue

        if '@' not in email:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Invalid email '{email}' sent to slask.") 

def Control_Phone(Webshop_Data):
    for index, row in Webshop_Data.iterrows():
        phone = str(row['Telefon']).strip()

        if pd.isnull(phone) or not phone:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Missing phone sent to slask.")
            continue

        phone = phone.replace(" ", "")

        if phone.startswith('46') and not phone.startswith('+46'):
            phone = '+' + phone

        if phone.startswith('+46') and len(phone) == 12 and phone[1:].isdigit():
            continue
        else:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Invalid phone '{phone}' sent to slask.")

def Control_Adress(Webshop_Data, Control_Data):


    for index, row in Webshop_Data.iterrows():
        adress = str(row['Full adress']).strip()

        if pd.isnull(adress) or not adress:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Missing adress sent to slask.")
            continue

        if not re.match(r"^[a-zA-Z0-9åäöÅÄÖ, ]+$", adress):
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Adress '{adress}' contains invalid characters. Sent to slask.")
            continue

        adress_match = adress.split(',')

        if len(adress_match) != 3:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Invalid address format. Sent to slask.")
            continue

        full_adress = adress_match[0].strip().capitalize()
        stad = adress_match[1].strip().capitalize()
        postnummer = adress_match[2].strip()

        if not re.match(r"^\d{5}$", postnummer):
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Invalid postnummer '{postnummer}'. Sent to slask.")
            continue

        match_found = False

        for _, control_row in Control_Data.iterrows():
            control_full_adress = control_row['Adress'].strip().capitalize()
            control_stad = control_row['Stad'].strip().capitalize()
            control_postnummer = str(control_row['Postnummer']).strip()

            if control_full_adress == full_adress and control_stad == stad and control_postnummer == postnummer:
                match_found = True
                break

        if not match_found:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Address '{adress}' not found in kontrollfilen. Sent to slask.")
        else:
            Webshop_Data.at[index, 'Full adress'] = full_adress
            Webshop_Data.at[index, 'Stad'] = stad
            Webshop_Data.at[index, 'Postnummer'] = postnummer


def Control_Customer_Registration(Webshop_Data):
    for index, row in Webshop_Data.iterrows():
        customer_registration = row['Kundregistrering']

        try:
            customer_registration = pd.to_datetime(customer_registration, format='%Y-%m-%d', errors='coerce')
            
            if pd.isnull(customer_registration):
                Webshop_Data.at[index, 'Slask'] = True
                print(f"Row {index}: Invalid registration date format sent to slask.")
                continue

            if customer_registration > pd.Timestamp.now():
                Webshop_Data.at[index, 'Slask'] = True
                print(f"Row {index}: Future registration date '{customer_registration}' sent to slask.")
        except Exception as e:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Error processing registration date '{customer_registration}' - {e}")

def Control_Produkt(Webshop_Data):

        for index, row in Webshop_Data.iterrows():
            product = str(row['Produkt']).strip()

            if pd.isnull(product) or not product:
                Webshop_Data.at[index, 'Slask'] = True
                print(f"Row {index}: Missing product sent to slask.")
                continue

            if not re.match(r"^[a-zA-Z0-9åäöÅÄÖ ]+$", product):
                Webshop_Data.at[index, 'Slask'] = True
                print(f"Row {index}: Product name '{product}' contains invalid characters. Sent to slask.")

def Control_Quantity(Webshop_Data):

    for index, row in Webshop_Data.iterrows():
        quantity = row['Kvantitet']

        if pd.isnull(quantity) or not quantity:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Missing quantity sent to slask.")
            continue

        try:
            quantity = int(quantity)
            if quantity <= 0:
                Webshop_Data.at[index, 'Slask'] = True
                print(f"Row {index}: Invalid quantity '{quantity}' (should be positive integer). Sent to slask.")
            continue
        except ValueError:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Invalid quantity '{quantity}' (not an integer). Sent to slask.")

def Control_PricePerUnit(Webshop_Data):

    for index, row in Webshop_Data.iterrows():
        priceperunit = row['Pris per enhet (kr)']

        if pd.isnull(priceperunit) or not priceperunit:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Missing price per unit sent to slask.")
            continue

        try:
            priceperunit = float(priceperunit)
            if priceperunit <= 0:
                Webshop_Data.at[index, 'Slask'] = True
                print(f"Row {index}: Invalid price per unit '{priceperunit}' (should be positive float). Sent to slask.")
                continue
        except ValueError:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Invalid price per unit '{priceperunit}' (not a valid float). Sent to slask.")

def Control_TotalPrice(Webshop_Data):

    for index, row in Webshop_Data.iterrows():
        totalprice = row['Total pris (kr)']

        if pd.isnull(totalprice) or totalprice == '':
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Missing total price sent to slask.")
            continue

        try:
            totalprice = float(totalprice)
            
            if totalprice <= 0:
                Webshop_Data.at[index, 'Slask'] = True
                print(f"Row {index}: Invalid total price '{totalprice}' (should be positive float). Sent to slask.")
                continue
        except ValueError:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Invalid total price '{totalprice}' (not a valid float). Sent to slask.")

def Control_TimeOfOrder(Webshop_Data):
    for index, row in Webshop_Data.iterrows():
        timeoforder = row['Ordertid']

        if pd.isnull(timeoforder) or timeoforder == '':
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Missing time of order sent to slask.")
            continue

        try:
            timeoforder_parsed = datetime.strptime(timeoforder, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            Webshop_Data.at[index, 'Slask'] = True
            print(f"Row {index}: Invalid time of order '{timeoforder}' (should be 'yyyy-mm-dd hh:mm:ss'). Sent to slask.")
            continue


def Finalize_Slask(Webshop_Data):

    Slask_data = Webshop_Data[Webshop_Data['Slask'] == True]
    valid_data = Webshop_Data[Webshop_Data['Slask'] == False]

    if not Slask_data.empty:
        if os.path.exists(Slaskfil):
            existing_Slask = pd.read_excel(Slaskfil)
            Slask_data = pd.concat([existing_Slask, Slask_data], ignore_index=True)

        Slask_data.to_excel(Slaskfil, index=False, engine='openpyxl')
        print(f"Slask file '{Slaskfil}' has been updated.")
    else:
        print("No new invalid rows found.")



Webshop_Data = pd.read_excel(Kunddata_Webbshop)
Webshop_Data['Slask'] = False
Control_Data = pd.read_excel(Control_Data)

Create_Database()
Create_Table_CustomerAdress()
Create_Table_Customer()
Create_Table_Purchase()
Create_Slask()
Control_Names(Webshop_Data)
Control_Birthdate(Webshop_Data)
Control_Email(Webshop_Data)
Control_Phone(Webshop_Data)
Control_Adress(Webshop_Data, Control_Data)
Control_Customer_Registration(Webshop_Data)
Control_Produkt(Webshop_Data)
Control_Quantity(Webshop_Data)
Control_PricePerUnit(Webshop_Data)
Control_TotalPrice(Webshop_Data)
Control_TimeOfOrder(Webshop_Data)
Finalize_Slask(Webshop_Data)

Cursor.close()
Connect.close()

 

del Webshop_Data

# Om du vill ta bort andra DataFrames du skapat:
# del slask_data, valid_data

# Rensa upp minnet
gc.collect()
print("DataFrames har tagits bort och minnet har rensats.")