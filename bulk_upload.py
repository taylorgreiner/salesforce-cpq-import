import pandas as pd
from simple_salesforce import Salesforce
from collections import defaultdict
from datetime import datetime
from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

# === 1. Salesforce Login ===
sf = Salesforce(
    username=os.getenv('SF_USERNAME'),
    password=os.getenv('SF_PASSWORD'),
    security_token=os.getenv('SF_SECURITY_TOKEN'),
    domain=os.getenv('SF_DOMAIN', 'login')  # Default to 'login' if not specified
)

# === 2. Load Excel File ===
excel_path = "your_excel_file.xlsx"
df = pd.read_excel(excel_path)

# Optional: clean date columns just in case
df["Contract Start Date"] = pd.to_datetime(df["Contract Start Date"], errors='coerce')
df["Term"] = pd.to_numeric(df["Term"], errors='coerce').fillna(12)

# === 3. Group Subscriptions by Order Number ===
orders = defaultdict(list)
for _, row in df.iterrows():
    order_number = row["Order Number"]
    orders[order_number].append(row)

# === 4. Loop Over Orders and Create Contracts + Subscriptions ===
for order_num, subs in orders.items():
    first_row = subs[0]
    
    try:
        account_id = first_row["Account Sales Force Id"]
        contract_start_date = first_row["Contract Start Date"]
        contract_term = int(first_row["Term"])

        # === Create Contract ===
        contract_response = sf.Contract.create({
            "AccountId": account_id,
            "StartDate": contract_start_date.strftime("%Y-%m-%d"),
            "ContractTerm": contract_term,
            "Status": "Draft",
            "Order_Number__c": order_num  # If custom field exists
        })
        contract_id = contract_response["id"]
        print(f"Created contract {contract_id} for order {order_num}")

        # === Create Subscriptions for Contract ===
        for row in subs:
            try:
                product_id = row["Product Sales Force Id"]
                quantity = float(row.get("Quantity", 1))
                mrc = float(row.get("Component Mrc", 0))
                product_name = row.get("Product", "Unknown Product")
                sub_start_date = row["Contract Start Date"]

                sub_data = {
                    "SBQQ__Contract__c": contract_id,
                    "SBQQ__Account__c": account_id,
                    "SBQQ__Product__c": product_id,
                    "SBQQ__Quantity__c": quantity,
                    "SBQQ__NetPrice__c": mrc,
                    "SBQQ__StartDate__c": sub_start_date.strftime("%Y-%m-%d"),
                    "Name": product_name
                }

                sf.SBQQ__Subscription__c.create(sub_data)
            except Exception as sub_err:
                print(f"Failed to create subscription for order {order_num}: {sub_err}")
        
        print(f"Finished creating {len(subs)} subscriptions for contract {contract_id}")

    except Exception as err:
        print(f"Error processing order {order_num}: {err}")
