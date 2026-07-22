import datetime
from xmlrpc import client
import numpy as np
import pandas as pd
from pymongo import MongoClient, ASCENDING, DESCENDING

def clean_and_convert_dates(val):
    """Safely converts Pandas Timestamps to native Python datetimes for MongoDB BSON compatibility."""
    if pd.isna(val):
        return None
    if isinstance(val, pd.Timestamp):
        return val.to_pydatetime()
    return val

def import_pipeline_to_mongodb(rent_file, link_file, mongo_uri="mongodb://localhost:27017/"):
    # 1. Establish MongoDB Connection
    print("Connecting to MongoDB...")
    client = MongoClient("mongodb+srv://wpsloperations:MHczqnxvMzav3Phv@wpsl-test.4jjnwbf.mongodb.net/")
    db = client["wolverine"]
    collection_name = "rent_transactions"
    
    # Re-initialize collection if it already exists (Wipes old data for a clean import)
    if collection_name in db.list_collection_names():
        print(f"Collection '{collection_name}' exists. Dropping for fresh import...")
        db[collection_name].drop()
        
    collection = db[collection_name]
    
    # 2. Load the Excel sheets
    print("Reading Excel files into DataFrames...")
    df_rent = pd.read_excel(rent_file, sheet_name='Rent_WOL')
    df_link = pd.read_excel(link_file, sheet_name='PO_Rent_Link')
    
    # Replace NaN values globally with Python None so MongoDB registers them as null
    df_rent = df_rent.replace({np.nan: None})
    df_link = df_link.replace({np.nan: None})
    
    # 3. Process and aggregate Purchase Orders
    print("Grouping and processing Purchase Orders...")
    po_dictionary = {}
    for _, row in df_link.iterrows():
        rent_id = row['Rent_ID']
        if rent_id is None:
            continue
            
        po_document = {
            "po_id": int(row['PO_ID']) if row['PO_ID'] is not None else None,
            "new_po_id": row['NEW_PO_ID'],
            "creation_date": clean_and_convert_dates(row['PO Creation_Date']),
            "payment_date": clean_and_convert_dates(row['Payment Date']),
            "allocated_amount": row['Payment amount in GBP'],
            "rent_commission": row['Rent_Commission'],
            "commission_amount": row['Commission_Amount'],
            "full_amount": row['Full_Amount'],
            "agent_commission": row['Agent_Commission'],
            "agent_id": row['Agent_ID'],
            "agent_name": row['Agent_Name'],
            "vat_percentage": row['VAT_Percentage'],
            "vat_amount": row['VAT_Amount'],
            "po_amount": row['PO_Amount'],
            "geo": {
                "town": row['town'],
                "city": row['city'],
                "post_code": row['post code']
            }
        }
        
        if rent_id not in po_dictionary:
            po_dictionary[rent_id] = []
        po_dictionary[rent_id].append(po_document)
        
    # 4. Construct Nested Document Structure
    print("Building unified document structures...")
    documents_to_insert = []
    
    for _, row in df_rent.iterrows():
        rent_id = row['ID']
        
        doc = {
            "_id": int(rent_id) if rent_id is not None else None,
            "source_ledger": "wolRentwol",
            "landlord": {
                "id": int(row['Landlord ID']) if row['Landlord ID'] is not None else None,
                "name": row['Landlord Name']
            },
            "property": {
                "id": int(row['Property id']) if row['Property id'] is not None else None,
                "flat_no": row['flat no'],
                "house_number": row['house number'],
                "street": row['street']
            },
            "schedule": {
                "start_date": clean_and_convert_dates(row['Start Date']),
                "end_date": clean_and_convert_dates(row['End Date']),
                "payment_date": clean_and_convert_dates(row['Payment Date'])
            },
            "financials": {
                "gross_rent_amount": row['Payment amount in GBP'],
                "rent_amount_for_agent": row['Rent_Amount_For_Agent'],
                "commission_from_qasim": row['Commission From Qasim'],
                "deductions_from_qasim": row['Deductions From Qasim'],
                "admin_commission_from_qasim": row['Admin Commission From Qasim'],
                "commission_from_ll": row['Commission From LL'],
                "deductions_from_ll": row['Deductions From LL'],
                "admin_commission_from_ll": row['Admin Commission From LL']
            },
            "status": {
                "has_rent_been_paid_to_ll": bool(row['Has Rent been paid to LL']) if row['Has Rent been paid to LL'] is not None else None,
                "payment_due": bool(row['Payment Due']) if row['Payment Due'] is not None else None,
                "date_paid_to_qasim": clean_and_convert_dates(row['Date Paid To Qasim']),
                "negotiated_date": clean_and_convert_dates(row['Negotiated Date']),
                "pay_in_slip_num": row['Pay_In_Slip_Num'],
                "checked": row['Checked'],
                "bank_date": clean_and_convert_dates(row['Bank Date']),
                "email_sent": bool(row['Email Sent']) if row['Email Sent'] is not None else None,
                "specific_method_of_payment": row['Specific Method Of Payment'],
                "bank_checked_other": bool(row['Bank Checked Other']) if row['Bank Checked Other'] is not None else None,
                "deposit": bool(row['Deposit?']) if row['Deposit?'] is not None else None,
                "never_po": bool(row['Never PO']) if row['Never PO'] is not None else None
            },
            "transaction_id": row['Transaction_ID'],
            "notes": row['Notes'],
            "query": row['Query'],
            "purchase_orders": po_dictionary.get(rent_id, [])  # Array-level embedding
        }
        documents_to_insert.append(doc)
        
    # 5. Bulk Insert into MongoDB
    if documents_to_insert:
        print(f"Inserting {len(documents_to_insert)} records into MongoDB...")
        result = collection.insert_many(documents_to_insert)
        print(f"Successfully inserted {len(result.inserted_ids)} documents.")
        
        # 6. Apply Indexes dynamically for query optimizations
        print("Generating collection indexes...")
        collection.create_index([("property.id", ASCENDING)])
        collection.create_index([("purchase_orders.new_po_id", ASCENDING)])
        collection.create_index([("schedule.start_date", DESCENDING)])
        print("Indexes successfully generated!")
    else:
        print("No records found to insert.")

    # Close connection
    client.close()

# --- Execution ---
if __name__ == "__main__":
    # Adjust files names and MongoDB URI string if necessary
    import_pipeline_to_mongodb(
        rent_file="wolRentwol.xlsx",
        link_file="PO_Rent_Link.xlsx",
        mongo_uri="mongodb://localhost:27017/"
    )