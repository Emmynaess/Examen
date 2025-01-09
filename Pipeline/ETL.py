import pyodbc
import os
import pandas as pd
import gc
import re
import phonenumbers
import datetime
from phonenumbers import NumberParseException, is_valid_number, format_number, PhoneNumberFormat
from decimal import Decimal, ROUND_HALF_UP

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
    'Packet Size=16384;' 
)

print('Koden startar')
current_dir = os.getcwd()
Kunddata_Webbshop = os.path.join(current_dir, 'kunddata_webbshop1.xlsx') 
Control_Data = os.path.join(current_dir, 'kunddata_adresser_kontroll1.xlsx')
Slaskfil = os.path.join(current_dir, 'slaskfil.xlsx')

Connect = pyodbc.connect(Connection_string, autocommit=True)
Cursor = Connect.cursor()

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
                FirstName NVARCHAR(20) NOT NULL,
                LastName NVARCHAR(20) NOT NULL,
                Phonenumber NVARCHAR(12) NOT NULL,
                Email NVARCHAR(50) UNIQUE NOT NULL,
                DateOfBirth DATE NOT NULL,
                StartOfMembership DATE NOT NULL
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

    return Slask_data

def Control_Names(Webshop_Data):

    if 'Slask' not in Webshop_Data.columns:
        Webshop_Data['Slask'] = False
    
    Slask_data = []
    invalid_chars = r"[#@\$%\^&\*!?=+]"

    for index, row in Webshop_Data.iterrows():
        customer_name = row.get('Kundnamn', None)

        if isinstance(customer_name, str):
            customer_name = re.sub(r"###\s*$", "", customer_name).strip()

        if pd.isnull(customer_name) or not isinstance(customer_name, str) or not customer_name.strip():
            Webshop_Data.at[index, 'Slask'] = True
            Slask_data.append(index)
            continue

        if re.search(invalid_chars, customer_name):
            Webshop_Data.at[index, 'Slask'] = True
            Slask_data.append(index)
            continue

        name_parts = customer_name.split()
        if len(name_parts) < 2:
            Webshop_Data.at[index, 'Slask'] = True
            Slask_data.append(index)
            continue

        first_name = name_parts[0].capitalize()
        last_name = ' '.join(part.capitalize() for part in name_parts[1:])

        if not first_name.strip() or not last_name.strip():
            Webshop_Data.at[index, 'Slask'] = True
            Slask_data.append(index)
        else:
            Webshop_Data.at[index, 'FirstName'] = first_name
            Webshop_Data.at[index, 'LastName'] = last_name

    if 'Kundnamn' in Webshop_Data.columns:
        Webshop_Data.drop(columns=['Kundnamn'], inplace=True)

    print(f"Control_Names: {len(Slask_data)} rows were sent to the slask file.")
    
    return Webshop_Data

def Control_Adress(Webshop_Data, Control_Data):

    try:
        Slask_data = []

        Webshop_Data['Full adress'] = Webshop_Data['Full adress'].astype(str).str.strip()
        adress_split = Webshop_Data['Full adress'].str.split(',', expand=True, n=2)

        if adress_split.shape[1] != 3:
            print("Step 3.1: Split did not produce 3 columns. Marking all rows as invalid.")
            Webshop_Data['Slask'] = True
            Slask_data.extend(Webshop_Data.index.tolist())
            return Webshop_Data

        adress_split.columns = ['StreetName', 'City', 'PostalCode']
        Webshop_Data[['StreetName', 'City', 'PostalCode']] = adress_split

        Webshop_Data['StreetName'] = Webshop_Data['StreetName'].str.strip().str.title()
        Webshop_Data['City'] = Webshop_Data['City'].str.strip().str.title()
        Webshop_Data['PostalCode'] = Webshop_Data['PostalCode'].str.strip().apply(lambda x: re.sub(r'[^0-9]', '', str(x)))

        rows_to_check = ~Webshop_Data['Slask']

        Control_Data['Adress'] = Control_Data['Adress'].str.strip().str.title()
        Control_Data['Stad'] = Control_Data['Stad'].str.strip().str.title()
        Control_Data['Postnummer'] = Control_Data['Postnummer'].astype(str).str.strip()

        Webshop_Data['Key'] = Webshop_Data['StreetName'] + '|' + Webshop_Data['City'] + '|' + Webshop_Data['PostalCode']
        Control_Data['Key'] = Control_Data['Adress'] + '|' + Control_Data['Stad'] + '|' + Control_Data['Postnummer']

        valid_keys = set(Control_Data['Key'])
        Webshop_Data.loc[rows_to_check, 'Slask'] = ~Webshop_Data.loc[rows_to_check, 'Key'].isin(valid_keys)

        Slask_data.extend(Webshop_Data.index[Webshop_Data['Slask']].tolist())

        Webshop_Data.drop(columns=['Full adress', 'Key'], inplace=True)

        print(f"Control_Adress: {len(Slask_data)} rows were sent to the slask file.")

    except Exception as e:
        print(f"An error occurred: {e}")
        Webshop_Data['Slask'] = True
        Slask_data.extend(Webshop_Data.index.tolist())

    return Webshop_Data

def Control_Birthdate(Webshop_Data):
    Slask_data = []

    for index, row in Webshop_Data.iterrows():
        if row.get('Slask', False):
            continue

        birthdate = row['Födelsedatum']
        if birthdate == "0000-00-00" or pd.isnull(birthdate) or not str(birthdate).strip():
            Webshop_Data.at[index, 'Slask'] = True
            Slask_data.append(index)
            continue
        try:
            birthdate = pd.to_datetime(birthdate, format='%Y-%m-%d', errors='raise').date()
            age = (pd.Timestamp.now().date() - birthdate).days // 365
            if age < 18:
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)
            else:
                Webshop_Data.at[index, 'DateOfBirth'] = birthdate 

        except Exception as e:
            Webshop_Data.at[index, 'Slask'] = True
            Slask_data.append(index)

            print(f"Error processing birthdate for row {index}: {e}")

    if 'Födelsedatum' in Webshop_Data.columns:
        Webshop_Data.drop(columns=['Födelsedatum'], inplace=True)
    print(f"Control_Birthdate: {len(Slask_data)} rows were sent to the slask file.")
    
    return Webshop_Data

def Control_Email(Webshop_Data):
    try:
        if 'Slask' not in Webshop_Data.columns:
            Webshop_Data['Slask'] = False

        Slask_data = []

        Webshop_Data['Email'] = Webshop_Data['Email'].astype(str)

        for index, row in Webshop_Data.iterrows():
            if row.get('Slask', False):
                continue

            email = row['Email']
            
            if pd.isnull(email) or email.strip() == '':
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)
                continue

            invalid_email = not bool(re.match(r'^[a-zA-Z0-9._%+-åäöÅÄÖ]+@[a-zA-Z0-9.-åäöÅÄÖ]+\.[a-zA-Z]{2,}$', email))
            if invalid_email:
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)

        print(f"Control_Email: {len(Slask_data)} rows were sent to the slask file.")

    except Exception as e:
        print(f"Fel vid kontroll av e-post: {e}")

def Control_Phone(Webshop_Data):

    try:
        if 'Slask' not in Webshop_Data.columns:
            Webshop_Data['Slask'] = False

        Slask_data = []

        if 'PhoneNumber' not in Webshop_Data.columns:
            Webshop_Data['PhoneNumber'] = None

        for index, row in Webshop_Data.iterrows():
            if row.get('Slask', False):
                continue

            phone = str(row.get('Telefon', '')).strip()

            if pd.isnull(phone) or not phone:
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)
                continue

            phone = phone.replace(" ", "")

            try:
                parsed_phone = phonenumbers.parse(phone, "SE")
                if is_valid_number(parsed_phone):
                    formatted_phone = format_number(parsed_phone, PhoneNumberFormat.INTERNATIONAL)
                    Webshop_Data.at[index, 'PhoneNumber'] = formatted_phone
                else:
                    Webshop_Data.at[index, 'Slask'] = True
                    Slask_data.append(index)
            except NumberParseException:
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)

        Webshop_Data['PhoneNumber'] = Webshop_Data['PhoneNumber'].str.replace(" ", "", regex=True)

        if 'Telefon' in Webshop_Data.columns:
            Webshop_Data.drop(columns=['Telefon'], inplace=True)

        print(f"Control_Phone: {len(Slask_data)} rows were sent to the slask file.")

    except Exception as e:
        print(f"An error occurred in Control_Phone: {e}")

    return Webshop_Data

def Control_Customer_Registration(Webshop_Data):
    try:
        if 'Slask' not in Webshop_Data.columns:
            Webshop_Data['Slask'] = False

        Slask_data = []

        for index, row in Webshop_Data.iterrows():
            if row.get('Slask', False):
                continue

            Customer_Registration = row.get('Kundregistrering')

            if (Customer_Registration == "0000-00-00" or 
                pd.isnull(Customer_Registration) or 
                not str(Customer_Registration).strip()):
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)
                continue

            try:
                Customer_Registration = pd.to_datetime(Customer_Registration, format='%Y-%m-%d', errors='raise').date()

                if Customer_Registration > datetime.date.today():
                    Webshop_Data.at[index, 'Slask'] = True
                    Slask_data.append(index)
                else:
                    Webshop_Data.at[index, 'Kundregistrering'] = Customer_Registration

            except Exception as e:
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)

        Webshop_Data.rename(columns={'Kundregistrering': 'Customer_Registration'}, inplace=True)

        print(f"Control_Customer_Registration: {len(Slask_data)} rows were sent to the slask file.")

    except Exception as e:
        print(f"An error occurred in Control_Customer_Registration: {e}")

    return Webshop_Data

def Control_Product(Webshop_Data):

    try:
        Slask_data = []

        if 'Produkt' in Webshop_Data.columns:
            Webshop_Data.rename(columns={'Produkt': 'Product'}, inplace=True)
        else:
            print("Column 'Produkt' not found. Skipping Control_Product.")
            return Webshop_Data

        for index, row in Webshop_Data.iterrows():

            if row.get('Slask', False):
                continue

            try:
                product = row['Product']

                if pd.isnull(product) or str(product).strip() == '' or len(str(product).strip()) > 100:
                    Webshop_Data.at[index, 'Slask'] = True
                    Slask_data.append(index)

            except Exception as e:
                print(f"Error processing row {index}: {e}")
                continue

        print(f"Control_Product: {len(Slask_data)} rows were sent to the slask file.")

        return Webshop_Data

    except Exception as e:
        print(f"An error occurred in Control_Product: {e}")

def Control_Quantity(Webshop_Data):

    Slask_data = []

    if 'Kvantitet' in Webshop_Data.columns:
        Webshop_Data.rename(columns={'Kvantitet': 'Quantity'}, inplace=True)
    else:
        return Webshop_Data

    for index, row in Webshop_Data.iterrows():

        if row.get('Slask', False):
            continue

        quantity = row['Quantity']

        try:

            if pd.isnull(quantity) or str(quantity).strip() == '':
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)
                continue

            quantity = int(quantity)
            if quantity <= 0:
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)
        except ValueError:
            Webshop_Data.at[index, 'Slask'] = True
            Slask_data.append(index)

    print(f"Control_Quantity: {len(Slask_data)} rows were sent to the slask file.")

    return Webshop_Data

def Control_Price_Per_Product(Webshop_Data):

    try:
        Slask_data = []

        if 'Pris per enhet (kr)' in Webshop_Data.columns:
            Webshop_Data.rename(columns={'Pris per enhet (kr)': 'PricePerProduct'}, inplace=True)
        else:
            return Webshop_Data

        for index, row in Webshop_Data.iterrows():

            if row.get('Slask', False):  # Hoppa över rader som redan är markerade som 'Slask'
                continue

            PricePerProduct = row.get('PricePerProduct', None)

            # Kontrollera om priset är tomt, icke-numeriskt eller <= 0
            if pd.isnull(PricePerProduct) or PricePerProduct == '' or PricePerProduct == 0:
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)
                continue

            try:
                if isinstance(PricePerProduct, str):
                    PricePerProduct = float(PricePerProduct.replace(',', '.'))

                if PricePerProduct <= 0:
                    Webshop_Data.at[index, 'Slask'] = True
                    Slask_data.append(index)
                    continue

                Webshop_Data.at[index, 'PricePerProduct'] = round(PricePerProduct, 2)

            except Exception:
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)

        print(f"Control_PricePerProduct: {len(Slask_data)} rows were sent to the slask file.")

    except Exception as e:
        print(f"Error in Control_Price_Per_Product: {e}")

    return Webshop_Data

def Control_Total_Price(Webshop_Data):

    try:
        Slask_data = []

        if 'Total pris (kr)' in Webshop_Data.columns:
            Webshop_Data.rename(columns={'Total pris (kr)': 'TotalPrice'}, inplace=True)
        else:
            return Webshop_Data

        for index, row in Webshop_Data.iterrows():
            totalprice = row['TotalPrice']

            if pd.isnull(totalprice) or str(totalprice).strip() == '':
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)
                continue

            try:

                if isinstance(totalprice, str):
                    totalprice = float(totalprice.replace(',', '.'))

                if totalprice <= 0:
                    Webshop_Data.at[index, 'Slask'] = True
                    Slask_data.append(index)
                    continue

                Webshop_Data.at[index, 'TotalPrice'] = round(totalprice, 2)

            except Exception:
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)

    except Exception as e:
        pass

    return Webshop_Data

def Control_Time_Of_Order(Webshop_Data):

    try:
        Slask_data = []

        if 'Ordertid' in Webshop_Data.columns:
            Webshop_Data.rename(columns={'Ordertid': 'TimeOfOrder'}, inplace=True)
        else:
            return Webshop_Data

        for index, row in Webshop_Data.iterrows():
            timeoforder = row['TimeOfOrder']

            if pd.isnull(timeoforder) or str(timeoforder).strip() == '' or "INVALID_" in str(timeoforder):
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)
                continue

            try:

                timeoforder_parsed = pd.to_datetime(timeoforder, errors='coerce', format='%Y-%m-%d %H:%M:%S')

                if pd.isnull(timeoforder_parsed):
                    Webshop_Data.at[index, 'Slask'] = True
                    Slask_data.append(index)
                    continue

                Webshop_Data.at[index, 'TimeOfOrder'] = timeoforder_parsed

            except Exception:
                Webshop_Data.at[index, 'Slask'] = True
                Slask_data.append(index)

        if 'Ordertid' in Webshop_Data.columns:
            Webshop_Data.drop(columns=['Ordertid'], inplace=True)

    except Exception:
        pass

    return Webshop_Data

def Finalize_Slask(Webshop_Data, Slaskfil):
    Slask_data = Webshop_Data[Webshop_Data['Slask'] == True]
    valid_data = Webshop_Data[Webshop_Data['Slask'] == False]

    if not Slask_data.empty:
        if os.path.exists(Slaskfil):
            existing_Slask = pd.read_excel(Slaskfil)
            Slask_data = pd.concat([existing_Slask, Slask_data], ignore_index=True)
        
        Slask_data.to_excel(Slaskfil, index=False, engine='openpyxl')
        print(f"Slask file '{Slaskfil}' has been updated with {len(Slask_data)} rows.")
    else:
        print("No new invalid rows found.")

    if not valid_data.empty:

        print(f"Valid data has been created with {len(valid_data)} rows.")
    else:
        print("No valid data to save.")

    return valid_data

def Insert_Customer_Data(valid_data, batch_size=100):

    try:

        conn = pyodbc.connect(Connection_string, autocommit=False) 
        cursor = conn.cursor()

        use_database_query = f"USE {Customer_Database}"
        cursor.execute(use_database_query)

        for i in range(0, len(valid_data), batch_size):
            batch = valid_data.iloc[i:i+batch_size]
            
            for index, row in batch.iterrows():
                try:
                    email = row.get('Email')

                    if row.get('Slask') is True:
                        continue

                    street_name = row.get('StreetName')
                    city = row.get('City')
                    postal_code = row.get('PostalCode')
                    first_name = row.get('FirstName')
                    last_name = row.get('LastName')
                    phone_number = row.get('PhoneNumber')
                    date_of_birth = row.get('DateOfBirth')
                    customer_registration = row.get('Customer_Registration')

                    address_check_query = f"""
                        SELECT CustomerAdressID 
                        FROM {CustomerAdress_Table} 
                        WHERE StreetName = ? AND City = ? AND PostalCode = ?
                    """
                    cursor.execute(address_check_query, street_name, city, postal_code)
                    address_id = cursor.fetchone()

                    if not address_id:

                        address_insert_query = f"""
                            INSERT INTO {CustomerAdress_Table} (StreetName, City, PostalCode) 
                            OUTPUT INSERTED.CustomerAdressID
                            VALUES (?, ?, ?)
                        """
                        cursor.execute(address_insert_query, street_name, city, postal_code)
                        result = cursor.fetchone()
                        if result:
                            address_id = result[0]
                        else:
                            continue
                    else:
                        address_id = address_id[0]

                    customer_check_query = f"""
                        SELECT CustomerID 
                        FROM {Customer_Table} 
                        WHERE FirstName = ? AND LastName = ? AND Email = ?
                    """
                    cursor.execute(customer_check_query, first_name, last_name, email)
                    customer_id = cursor.fetchone()

                    if not customer_id:

                        customer_insert_query = f"""
                            INSERT INTO {Customer_Table} (CustomerAdressID, FirstName, LastName, Phonenumber, Email, DateOfBirth, StartOfMembership)
                            OUTPUT INSERTED.CustomerID
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                        """
                        cursor.execute(customer_insert_query, address_id, first_name, last_name, phone_number, email, date_of_birth, customer_registration)
                        result = cursor.fetchone()
                        if result:
                            customer_id = result[0]
                        else:
                            continue
                    else:
                        customer_id = customer_id[0]

                    if 'Product' in row and 'Quantity' in row and 'PricePerProduct' in row and 'TotalPrice' in row:
                        product = row.get('Product')
                        quantity = row.get('Quantity')
                        price_per_product = row.get('PricePerProduct')
                        total_price = row.get('TotalPrice')
                        time_of_purchase = row.get('OrderTime') if 'OrderTime' in row else row.get('TimeOfOrder')

                        if product and quantity and price_per_product and total_price and time_of_purchase:
                            purchase_check_query = f"""
                                SELECT PurchaseID 
                                FROM {Purchase_Table} 
                                WHERE CustomerID = ? AND Product = ? AND TimeOfPurchase = ?
                            """
                            cursor.execute(purchase_check_query, customer_id, product, time_of_purchase)
                            purchase_id = cursor.fetchone()

                            if not purchase_id:

                                purchase_insert_query = f"""
                                    INSERT INTO {Purchase_Table} (CustomerID, Product, Quantity, PricePerProduct, TotalPrice, TimeOfPurchase)
                                    OUTPUT INSERTED.PurchaseID
                                    VALUES (?, ?, ?, ?, ?, ?)
                                """
                                cursor.execute(purchase_insert_query, customer_id, product, quantity, price_per_product, total_price, time_of_purchase)
                except Exception as row_error:
                    print(f"Error processing row {index}: {row_error}")
                    continue

            conn.commit()
            print(f"Batch {i // batch_size + 1} inserted successfully!")

        print("All data has been successfully inserted.")

    except Exception as e:
        print(f"Error occurred during data insertion: {e}")
        conn.rollback()
    finally:
        conn.close()

Webshop_Data = pd.read_excel(Kunddata_Webbshop, engine='openpyxl')
Webshop_Data['Slask'] = False
Control_Data = pd.read_excel(Control_Data)
print(f"Antal rader i Webshop_Data: {len(Webshop_Data)}")

Create_Database()
Create_Table_CustomerAdress()
Create_Table_Customer()
Create_Table_Purchase()
Create_Slask()
Control_Names(Webshop_Data)
Control_Adress(Webshop_Data, Control_Data)
Control_Birthdate(Webshop_Data)
Control_Email(Webshop_Data)
Control_Phone(Webshop_Data)
Control_Customer_Registration(Webshop_Data)
Control_Product(Webshop_Data)
Control_Quantity(Webshop_Data)
Control_Price_Per_Product(Webshop_Data)
Control_Total_Price(Webshop_Data)
Control_Time_Of_Order(Webshop_Data)
valid_data = Finalize_Slask(Webshop_Data, Slaskfil)
Insert_Customer_Data(valid_data, batch_size=100)

del Webshop_Data

gc.collect()
print("DataFrames har tagits bort och minnet har rensats.")