import pandas as pd
from pymongo import MongoClient, UpdateOne
import datetime

# 1. Connect to your database
client = MongoClient("mongodb+srv://wpsloperations:MHczqnxvMzav3Phv@wpsl-test.4jjnwbf.mongodb.net/")
db = client["wolverine"]
collection = db["properties"] # Change to your collection name

# 2. Read the excel sheet
df = pd.read_excel('PreviousTenancy.xlsx')
 
def format_date(val):
    if pd.isna(val):
        return None
    return pd.to_datetime(val)

def format_decimal(val):
    if pd.isna(val):
        return 0.0
    return float(val)

def format_str(val):
    if pd.isna(val):
        return None
    return str(val).strip()

# 3. Compile updates grouped by Property ID
bulk_operations = []

# Group entries by property to push all records in one go per document
for prop_id, group in df.groupby('Property ID'):
    past_records = []
    
    for _, row in group.iterrows():
        tenancy_object = {
            "client_id": int(row['Client_ID']) if not pd.isna(row['Client_ID']) else None,
            "tenant_name": format_str(row['Client_Name']),
            "phone_number": format_str(row['Client_Phone_Number']),
            "email_address": format_str(row['Client_Email_Address']),
            "start_date": format_date(row['Start Date']),
            "end_date": format_date(row['End Date']),
            "deposit_amount": format_decimal(row['Deposit Amount']),
            "deposit_returned_date": format_date(row['Deposit Returned Date']),
            "forwarding_address": format_str(row['Client_Forwarding_Address']),
            "notes": format_str(row['PreviousTenancy.Notes'])
        }
        past_records.append(tenancy_object)
    
    # Using $set to completely override or $push with $each to append
    bulk_operations.append(
        UpdateOne(
            { "property_id": int(prop_id) },
            { "$set": { "tenancy_records.past_tenancies": past_records } }
        )
    )

# 4. Execute Bulk Write for max performance
if bulk_operations:
    result = collection.bulk_write(bulk_operations)
    print(f"Successfully matched {result.matched_count} properties and updated past tenancies!")