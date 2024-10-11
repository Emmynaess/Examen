import pandas as pd
import random
from faker import Faker
 
# Skapa en Faker-instans med svenska inställningar
fake = Faker('sv_SE')
 
# Definiera en lista med slumpmässiga varor
produkter = [
    "Laptop", "Mobiltelefon", "Surfplatta", "Hörlurar", "Smartklocka",
    "Kamera", "TV", "Högtalare", "Spelkonsol", "Skrivare", "Router"
]
 
# Funktion för att generera kunddata
def generera_kunddata(antal_rader):
    data = []
 
    for _ in range(antal_rader):
        namn = fake.name()
        email = fake.email()
        stad = fake.city()
        postnummer = fake.postcode()
        telefon = fake.phone_number()
 
        produkt = random.choice(produkter)
        kvantitet = random.randint(1, 5)
        pris_per_enhet = random.uniform(100, 15000)  # Pris mellan 100 kr och 15000 kr
        total_pris = round(pris_per_enhet * kvantitet, 2)
 
        data.append({
            "Kundnamn": namn,
            "Email": email,
            "Telefon": telefon,
            "Stad": stad,
            "Postnummer": postnummer,
            "Produkt": produkt,
            "Kvantitet": kvantitet,
            "Pris per enhet (kr)": round(pris_per_enhet, 2),
            "Total pris (kr)": total_pris
        })
 
    return pd.DataFrame(data)
 
# Generera 100 000 rader kunddata
kunddata = generera_kunddata(100000)
 
# Spara data till en CSV-fil
kunddata.to_csv("kunddata_webbshop.csv", index=False, encoding='utf-8-sig')