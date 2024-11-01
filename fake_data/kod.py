"""import pandas as pd
import random
from faker import Faker
import re
from datetime import datetime, timedelta

fake = Faker('sv_SE')

produkter = [
    "Laptop", "Mobiltelefon", "Surfplatta", "Hörlurar", "Smartklocka",
    "Kamera", "TV", "Högtalare", "Spelkonsol", "Skrivare", "Router"
]

def replace_bokstaver(s):
    s = s.lower()
    s = s.replace('å', 'a').replace('ä', 'a').replace('ö', 'o')
    s = s.replace('é', 'e').replace('ü', 'u')
    s = re.sub(r'[^a-z]', '', s)
    return s

def generera_telefonnummer():
    prefix = random.choice(['+4670', '+4672', '+4673', '+4676', '+4679'])
    nummer = ''.join([str(random.randint(0, 9)) for _ in range(7)])
    telefonnummer = prefix + nummer
    return telefonnummer

def generera_kunddata(antal_rader):
    data = []

    registration_start_date = datetime.now() - timedelta(days=1094)  # 3 år
    registration_end_date = datetime.now() - timedelta(days=1)  # Igår

    order_start_date = datetime.now() - timedelta(days=365)
    order_end_date = datetime.now()

    for _ in range(antal_rader):
        first_name = fake.first_name()
        last_name = fake.last_name()
        namn = f"{first_name} {last_name}"
        
        first_name_clean = replace_bokstaver(first_name)
        last_name_clean = replace_bokstaver(last_name)
        
        email_domain = random.choice(['gmail.com', 'email.com', 'email.se', 'outlook.com', 'hotmail.com', 'live.se'])
        email = f"{first_name_clean}.{last_name_clean}@{email_domain}"
        
        telefon = generera_telefonnummer()

        gatuadress = fake.street_address()
        stad = fake.city()
        postnummer = fake.postcode()

        full_adress = f"{gatuadress}, {stad}, {postnummer}"

        produkt = random.choice(produkter)
        kvantitet = random.randint(1, 5)
        pris_per_enhet = round(random.uniform(100, 15000), 2)
        total_pris = round(pris_per_enhet * kvantitet, 2)

        registration_timestamp = registration_start_date + (registration_end_date - registration_start_date) * random.random()
        kundregistrering = registration_timestamp.strftime('%Y-%m-%d')

        order_timestamp = registration_timestamp + (order_end_date - registration_timestamp) * random.random()
        ordertid = order_timestamp.strftime('%Y-%m-%d %H:%M:%S')


        data.append({
            "Kundnamn": namn,
            "Email": email,
            "Telefon": telefon,
            "Full adress": full_adress,
            "Kundregistrering": kundregistrering,
            "Produkt": produkt,
            "Kvantitet": kvantitet,
            "Pris per enhet (kr)": pris_per_enhet,
            "Total pris (kr)": total_pris,
            "Ordertid": ordertid
        })

    return pd.DataFrame(data)

kunddata = generera_kunddata(500)

kunddata.to_excel("kunddata_webbshop.xlsx", index=False, engine='openpyxl')"""




"""import pandas as pd
import random
from faker import Faker
import re

# Skapa en Faker-instans med svenska inställningar
fake = Faker('sv_SE')

# Definiera en lista med slumpmässiga varor
produkter = [
    "Laptop", "Mobiltelefon", "Surfplatta", "Hörlurar", "Smartklocka",
    "Kamera", "TV", "Högtalare", "Spelkonsol", "Skrivare", "Router"
]

def replace_bokstaver(s):
    s = s.lower()
    s = s.replace('å', 'a').replace('ä', 'a').replace('ö', 'o')
    s = s.replace('é', 'e').replace('ü', 'u')
    s = re.sub(r'[^a-z]', '', s)
    return s

def generera_telefonnummer():
    prefix = random.choice(['+4670', '+4672', '+4673', '+4676', '+4679'])
    nummer = ''.join([str(random.randint(0, 9)) for _ in range(7)])
    telefonnummer = prefix + nummer
    return telefonnummer

# Funktion för att generera kunddata
def generera_kunddata(antal_rader):
    data = []

    for _ in range(antal_rader):
        # Generera namn och e-post
        first_name = fake.first_name()
        last_name = fake.last_name()
        namn = f"{first_name} {last_name}"
        
        # Rensa bokstäver och skapa e-postadress
        first_name_clean = replace_bokstaver(first_name)
        last_name_clean = replace_bokstaver(last_name)
        email_domain = random.choice(['gmail.com', 'email.com', 'email.se', 'outlook.com'])
        email = f"{first_name_clean}.{last_name_clean}@{email_domain}"
        
        # Generera telefonnummer
        telefon = generera_telefonnummer()
        
        # Generera adress, stad och postnummer
        gatuadress = fake.street_address()  # Hämta en gatuadress
        stad = fake.city()  # Hämta en stad
        postnummer = fake.postcode()  # Hämta ett postnummer

        # Kombinera gatuadress, stad och postnummer i en kolumn
        full_adress = f"{gatuadress}, {stad}, {postnummer}"

        # Generera produktinformation
        produkt = random.choice(produkter)
        kvantitet = random.randint(1, 5)
        pris_per_enhet = round(random.uniform(100, 15000), 2)  # Pris mellan 100 kr och 15000 kr
        total_pris = round(pris_per_enhet * kvantitet, 2)

        # Lägg till en rad med kunddata
        data.append({
            "Kundnamn": namn,
            "Email": email,
            "Telefon": telefon,
            "Full adress": full_adress,  # Lägg till full adress i en kolumn
            "Produkt": produkt,
            "Kvantitet": kvantitet,
            "Pris per enhet (kr)": pris_per_enhet,
            "Total pris (kr)": total_pris
        })

    return pd.DataFrame(data)

# Generera 500,000 rader kunddata
kunddata = generera_kunddata(500000)

# Spara data till en Excel-fil
kunddata.to_excel("kunddata_webbshop.xlsx", index=False, engine='openpyxl')"""