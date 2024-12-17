import pandas as pd
from faker import Faker
import random
import requests
import numpy as np
from config import AZURE_MAPS_API_KEY
import time
import logging
import unicodedata
from datetime import datetime
import os

fake = Faker("sv_SE")
logging.basicConfig(level=logging.INFO)

def load_products_from_csv(filename):

    try:
        filepath = os.path.join(os.path.dirname(__file__), filename)
        products_df = pd.read_csv(filepath)
        products = products_df.to_dict(orient="records")  # Konvertera till lista med dictionaries
        return products
    except Exception as e:
        logging.error(f"Failed to load products from {filename}: {e}")
        return []

def generate_coordinates():
    lat = random.uniform(58.0, 63.0)
    lon = random.uniform(14.0, 20.0)
    return lat, lon

def get_address_from_coordinates(lat, lon, api_key):
    url = "https://atlas.microsoft.com/search/address/reverse/json"
    params = {
        'api-version': '1.0',
        'subscription-key': api_key,
        'query': f"{lat},{lon}",
        'language': 'sv-SE',
        'countrySet': 'SE'
    }

    response = requests.get(url, params=params, timeout=10)
    response.raise_for_status()
    data = response.json()
    addresses = data.get('addresses', [])
    if addresses:
        address_info = addresses[0].get('address', {})
        street = address_info.get('streetName', 'No Street')
        house_number = address_info.get('streetNumber', 'No Number')
        postcode = address_info.get('postalCode', 'No Postcode')
        city = address_info.get('municipalitySubdivision', 'No City')
        municipality = address_info.get('municipality', 'No Municipality')
        full_street = f"{street} {house_number}".strip()
        return full_street, postcode, city, municipality
    else:
        return None, None, None, None

def clean_string(s):
    s = unicodedata.normalize('NFKD', s)
    s = s.encode('ascii', 'ignore').decode('ascii')
    s = s.lower()
    s = s.replace(' ', '')
    return s

def generate_swedish_phone_number():
    area_code = random.choice(["70", "72", "73", "76"])
    first_part = random.randint(100, 999)
    second_part = random.randint(10, 99)
    third_part = random.randint(10, 99)
    return f"+46 {area_code}-{first_part} {second_part} {third_part}"

def generate_data(rows=10, max_retries=10):
    domains = ["hotmail.com", "gmail.com", "outlook.com", "live.com", "icloud.com"]
    products = load_products_from_csv("products.csv")

    data = []

    for i in range(rows):
        first_name = fake.first_name()
        last_name = fake.last_name()
        birthdate = fake.date_of_birth(minimum_age=18, maximum_age=90).strftime('%Y-%m-%d')
        phone = generate_swedish_phone_number()
        email = f"{clean_string(first_name)}.{clean_string(last_name)}@{random.choice(domains)}"
        customer_category = random.choice(["Private", "Business"])

        
        for attempt in range(max_retries):
            lat, lon = generate_coordinates()
            street, postcode, city, municipality = get_address_from_coordinates(lat, lon, AZURE_MAPS_API_KEY)
            if street != "No Street" and postcode != "No Postcode" and city != "No City" and municipality != "No Municipality":
                break
            else:
                logging.warning(f"Attempt {attempt+1}: Could not retrieve a valid address. Retrying...")
        else:
            logging.error(f"Using fallback address after {max_retries} attempts.")
            street, postcode, city, municipality = "Fallback Street", "Fallback Postcode", "Fallback City", "Fallback Municipality"

        
        purchase_count = random.randint(1, 5)
        for _ in range(purchase_count):
            purchase_date = fake.date_between(start_date='-3y', end_date='today').strftime('%Y-%m-%d')
            product = random.choice(products)
            quantity = random.randint(1, 5)
            total_amount = product["price"] * quantity

            data.append({
                "First Name": first_name,
                "Last Name": last_name,
                "Birthdate": birthdate,
                "Phone": phone,
                "Email": email,
                "Customer Category": customer_category,
                "Street": street,
                "Postcode": postcode,
                "City": city,
                "Municipality": municipality,
                "Purchase Date": purchase_date,
                "Product": product["product"],
                "Quantity": quantity,
                "Price per Item": product["price"],
                "Total Amount": total_amount
            })

    df = pd.DataFrame(data)
    return df

def save_to_excel(df, filename):
    df.to_excel(filename, index=False)
    print(f"Excel file '{filename}' has been created.")

customer_data = generate_data(10)
save_to_excel(customer_data, "customer_data.xlsx")
