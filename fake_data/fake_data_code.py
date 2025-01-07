import pandas as pd
import random
from faker import Faker
import re
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import openpyxl

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

def introduce_errors(row):

    if random.random() < 0.1:
        row['Email'] = None

    if random.random() < 0.05:
        row['Kundnamn'] = row['Kundnamn'] + ' ###'

    if random.random() < 0.1:
        row['Telefon'] = '12' + ''.join([str(random.randint(0, 9)) for _ in range(random.choice([5, 15]))])

    if random.random() < 0.05:
        row['Total pris (kr)'] = -abs(row['Total pris (kr)'])

    if random.random() < 0.05:
        row['Ordertid'] = ''

    if random.random() < 0.02:
        row['Full adress'] = row['Full adress'] + ', ###'

    if random.random() < 0.05:
        row['Födelsedatum'] = '0000-00-00'

    if random.random() < 0.1:
        row['Produkt'] = ''

    if random.random() < 0.05:
        row['Kvantitet'] = -random.randint(1, 5)

    if random.random() < 0.1:
        row['Pris per enhet (kr)'] = -round(random.uniform(100, 500000), 2)

    if random.random() < 0.03:
        row['Kundregistrering'] = ''

    return row

def generera_kunddata(antal_rader):
    
    data = []

    registration_start_date = datetime.now() - timedelta(days=1094)
    registration_end_date = datetime.now() - timedelta(days=1)

    order_start_date = datetime.now() - timedelta(days=365)
    order_end_date = datetime.now()

    for _ in range(antal_rader):
        first_name = fake.first_name()
        last_name = fake.last_name()
        namn = f"{first_name} {last_name}"
        
        first_name_clean = replace_bokstaver(first_name)
        last_name_clean = replace_bokstaver(last_name)
        
        email_domain = random.choice(['Hotmail.se','Hotmail.com','live.se','gmail.se','gmail.com', 'email.com', 'email.se', 'outlook.com'])
        email = f"{first_name_clean}.{last_name_clean}@{email_domain}"
        
        telefon = generera_telefonnummer()

        gatuadress = fake.street_address()
        stad = fake.city()
        postnummer = fake.postcode()

        full_adress = f"{gatuadress}, {stad}, {postnummer}"

        fodelsedatum = fake.date_of_birth(minimum_age=18, maximum_age=65)
        fodelsedatum_str = fodelsedatum.strftime('%Y-%m-%d')

        registration_date = fake.date_between_dates(
            date_start=fodelsedatum + relativedelta(years=18),
            date_end=registration_end_date.date()
        )

        registration_timestamp = datetime.combine(registration_date, datetime.min.time())
        kundregistrering = registration_timestamp.strftime('%Y-%m-%d')

        time_diff = order_end_date - registration_timestamp
        random_seconds = random.uniform(0, time_diff.total_seconds())
        order_timestamp = registration_timestamp + timedelta(seconds=random_seconds)
        ordertid = order_timestamp.strftime('%Y-%m-%d %H:%M:%S')

        produkt = random.choice(produkter)
        kvantitet = random.randint(1, 5)
        pris_per_enhet = round(random.uniform(100, 500000), 2)
        total_pris = round(pris_per_enhet * kvantitet, 2)

        row = {
            "Kundnamn": namn,
            "Födelsedatum": fodelsedatum_str,
            "Email": email,
            "Telefon": telefon,
            "Full adress": full_adress,
            "Kundregistrering": kundregistrering,
            "Produkt": produkt,
            "Kvantitet": kvantitet,
            "Pris per enhet (kr)": pris_per_enhet,
            "Total pris (kr)": total_pris,
            "Ordertid": ordertid
        }

        data.append(row)

    return pd.DataFrame(data)

kunddata = generera_kunddata(500000)

adress_data = kunddata[['Full adress']].copy()
adress_data[['Adress', 'Stad', 'Postnummer']] = adress_data['Full adress'].str.extract(r'(.+),\s*(.+),\s*(\d+)$')
adress_data = adress_data.drop(columns=['Full adress'])

adress_data.to_excel("kunddata_adresser_kontroll.xlsx", index=False, engine='openpyxl')

kunddata = kunddata.apply(introduce_errors, axis=1)

kunddata.to_excel("kunddata_webbshop.xlsx", index=False, engine='openpyxl')

print("Kunddata och kontrollfil skapades. Fel har introducerats i 'kunddata_webbshop.xlsx'.")
