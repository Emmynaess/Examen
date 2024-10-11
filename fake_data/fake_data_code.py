import pandas as pd
import random
from faker import Faker
import re

# addressinfo hamnar i samma kolumn i excel
# lägga till adress och gatnummer
# Lägg till så att det hamnar i excel och inte csv fil

# ändrat så det bara är svenska telefonnummer
# ändrat så att email stämmer överens med kundnamn + tagit bort å ä ö ur emailadressen

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

    for _ in range(antal_rader):
        first_name = fake.first_name()
        last_name = fake.last_name()
        namn = f"{first_name} {last_name}"

        first_name_clean = replace_bokstaver(first_name)
        last_name_clean = replace_bokstaver(last_name)
        email_domain = random.choice(['gmail.com', 'email.com', 'email.se', 'outlook.com'])
        email = f"{first_name_clean}.{last_name_clean}@{email_domain}"

        telefon = generera_telefonnummer()
        adress = fake.street_address().replace('\n', ', ')
        stad = fake.city()
        postnummer = fake.postcode()

        produkt = random.choice(produkter)
        kvantitet = random.randint(1, 5)
        pris_per_enhet = round(random.uniform(100, 15000), 2)
        total_pris = round(pris_per_enhet * kvantitet, 2)

        data.append({
            "Kundnamn": namn,
            "Email": email,
            "Telefon": telefon,
            "Adress": adress,
            "Stad": stad,
            "Postnummer": postnummer,
            "Produkt": produkt,
            "Kvantitet": kvantitet,
            "Pris per enhet (kr)": pris_per_enhet,
            "Total pris (kr)": total_pris
        })

    return pd.DataFrame(data)

# börjar med hundra nu i testläget
kunddata = generera_kunddata(100)

kunddata.to_excel("kunddata_webbshop.xlsx", index=False, engine='openpyxl')