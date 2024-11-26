import pandas as pd
import random
from faker import Faker
import re
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import openpyxl

fake = Faker('sv_SE')

products = [
    "Laptop", "Mobile Phone", "Tablet", "Headphones", "Smartwatch",
    "Camera", "TV", "Speaker", "Game Console", "Printer", "Router"
]

def replace_letters(s):
    s = s.lower()
    s = s.replace('å', 'a').replace('ä', 'a').replace('ö', 'o')
    s = s.replace('é', 'e').replace('ü', 'u')
    s = re.sub(r'[^a-z]', '', s)
    return s

def generate_phone_number():
    prefix = random.choice(['+4670', '+4672', '+4673', '+4676', '+4679'])
    number = ''.join([str(random.randint(0, 9)) for _ in range(7)])
    phone_number = prefix + number
    return phone_number

def introduce_errors(row):
    """Introduce random errors into the data."""
    if random.random() < 0.1:  
        row['Email'] = None
    if random.random() < 0.05:  
        row['Customer Name'] = row['Customer Name'] + ' @#$'
    if random.random() < 0.1:  
        row['Phone'] = '12345'
    if random.random() < 0.05:  
        row['Total Price (kr)'] = -row['Total Price (kr)']
    return row

def generate_customer_data(row_count):
    data = []

    registration_start_date = datetime.now() - timedelta(days=1094)  # 3 years ago
    registration_end_date = datetime.now() - timedelta(days=1)  # Yesterday

    order_start_date = datetime.now() - timedelta(days=365)
    order_end_date = datetime.now()

    for _ in range(row_count):
        first_name = fake.first_name()
        last_name = fake.last_name()
        name = f"{first_name} {last_name}"
        
        first_name_clean = replace_letters(first_name)
        last_name_clean = replace_letters(last_name)
        
        email_domain = random.choice(['gmail.com', 'email.com', 'email.se', 'outlook.com', 'hotmail.com', 'live.se'])
        email = f"{first_name_clean}.{last_name_clean}@{email_domain}"
        
        phone = generate_phone_number()

        street_address = fake.street_address()
        city = fake.city()
        postal_code = fake.postcode()

        full_address = f"{street_address}, {city}, {postal_code}"

        birthdate = fake.date_of_birth(minimum_age=18, maximum_age=65)
        birthdate_str = birthdate.strftime('%Y-%m-%d')

        registration_date = fake.date_between_dates(
            date_start=birthdate + relativedelta(years=18),
            date_end=registration_end_date.date()
        )

        registration_timestamp = datetime.combine(registration_date, datetime.min.time())
        customer_registration = registration_timestamp.strftime('%Y-%m-%d')

        time_diff = order_end_date - registration_timestamp
        random_seconds = random.uniform(0, time_diff.total_seconds())
        order_timestamp = registration_timestamp + timedelta(seconds=random_seconds)
        order_time = order_timestamp.strftime('%Y-%m-%d %H:%M:%S')

        product = random.choice(products)
        quantity = random.randint(1, 5)
        price_per_unit = round(random.uniform(100, 15000), 2)
        total_price = round(price_per_unit * quantity, 2)

        row = {
            "Customer Name": name,
            "Birthdate": birthdate_str,
            "Email": email,
            "Phone": phone,
            "Full Address": full_address,
            "Customer Registration": customer_registration,
            "Product": product,
            "Quantity": quantity,
            "Price per Unit (kr)": price_per_unit,
            "Total Price (kr)": total_price,
            "Order Time": order_time
        }

        row = introduce_errors(row)
        data.append(row)

    return pd.DataFrame(data)

customer_data = generate_customer_data(150)

customer_data.to_excel("customer_data.xlsx", index=False, engine='openpyxl')