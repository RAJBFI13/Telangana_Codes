import pandas as pd
import re

# Load the Excel file
try:
    transaction_data = pd.read_excel("C:\\Users\\pc\\Desktop\\All Scripts\\Telangana\\Data\\Excel Data\\SECUNDERABAD.xlsx", header=0)
    transaction_data = transaction_data.astype(str)
except Exception as e:
    print(f"Error loading Excel file: {str(e)}")
    exit(1)

# Define regular expressions for data extraction
document_type_pattern = r'(\d+)[\r\n]+([A-Za-z ]+)'
locality_pattern = r'\/([\w\s-]+)-\d+\s+W-B:'
house_pattern = r'HOUSE:\s([\w/-]+)'
apartment_pattern = r'APARTMENT:\s*([^\d@]+) FLAT:'
flat_pattern = r'FLAT:\s(\d+)'
built_pattern = r'BUILT:\s([\d.]+)'
floor_pattern = r'(\d+)(?:ST|ND|RD|TH)\s+FLR'
floor_pattern02 = r'(\S+)\s+FLOORBoundires'
village = "Hyderabad"
Status = "ready"
city = "650c1e6c3ddfec1666f779bc"
state = "650c1dea3ddfec1666f779ba"

# Commercial keywords
Commercial = ["SHOP", "OFFICE"]

# KEY ABBREVIATIONS OF SELLER
Buyer = ["CL", "MR", "DR", "LR", "RR", "PL"]

# KEY ABBREVIATIONS OF SELLER
Seller = ["EX", "ME", "DE", "LE", "RE", "AY"]

#COLUMNS TO CONVERT FROM STRING TO INTEGER
columns_to_convert = ["DocumentNo", "S.No.", "MarketPrice", "Construction Value", "FlatNo"]

#BUILDING NAME SUPPOSED TO BE REPLACED
Build_Name = ["X","XX","XXX","XXXX","XXXXX"]
# Convert columns to string
transaction_data = transaction_data.astype(str)

# Lists of empty keys for entities and building data
entity_columns = ["PurchaserEmail", "SellerEmail", "PurchaserContact", "SellerContact"]

#EMPTY KEYS OF TRANSACTION DATA
transaction_empty_keys = ["PurchaserEmail", "SellerEmail", "PurchaserContact", "SellerContact",
                       "SROName", "BlockNo", "PinCode", "Others", "AadharNO", "PAN_No", "AgreementNo", "Time",
                       "DocumentSerialNo", "DHCfeesOrDocumentHandlingCharges", "RegistrationFees", "Age",
                       "BazarMulyaOrMarketRate", "MobdalaOrConsideration", "Bharlele_Mudrank_ShulkhOr_Stamp_Duty_Paid",
                       "MTR", "rate", "Corporate_Identification_number_or_CIN Parking_Information",
                       "License_Period Lock_In_Period Fit_our_eriod", "Escalation_in_Licensee_fees",
                       "CAM_Or_Common_Area_Maintenance", "Security_Deposit", "SecondaryRegistrar", "Compensation",
                       "PropertyDescription", "SROCode"]

#BUILDING DATA EMPTY KEYS
building_empty_keys = ["BlockNo", "PinCode", "Others", "buildingAge", "totalFloor", "totalFlat", "totalRegistryTransaction",
                       "developer", "lat", "long", "rent/Compensation", "assetsId", "unitCondition", "unitOccupancyStatus", "efficiency",
                       "loading", "commercialsChargeable", "commercialsCarpet", "CAMChargeable", "CAMCarpet", "propTax", "lockInMonths",
                       "securityDepositMonths", "commonCafeteria", "powerBackup", "aCType", "remarks", "parkingInformation", "commonAreaMaintenance",
                       "carpetSqft", "chargeableSqft", "floorPlateCarpet", "floorplateChargeable", "area"]


# Remove abbreviations in "Reg.Date", "Exe.Date", "Pres.Date"
transaction_data["Reg.Date"] = transaction_data["Reg.Date"].str.replace(r'\(R\)\s*', '', regex=True)
transaction_data["Exe.Date"] = transaction_data["Exe.Date"].str.replace(r'\(E\)\s*', '', regex=True)
transaction_data["Pres.Date"] = transaction_data["Pres.Date"].str.replace(r'\(P\)\s*', '', regex=True)

# TO REMOVE ALL THE EXTRA WHITESPACES FROM THE DATA FRAME
def strip_whitespace(cell):
    if isinstance(cell, str):
        return cell.strip()
    else:
        return cell

# Function to extract Seller and Buyer Name
def extract_ex_cl(text):
    try:
        ex_cl_pattern = r'(\d+\.\s*\([^)]+\)[^\d]*)'
        ex_cl_entries = re.findall(ex_cl_pattern, text)

        ex_entries = [re.sub(r'\d+|\([^)]+\)|\.', '', entry).strip() for entry in ex_cl_entries for buy in Buyer if buy in entry]
        cl_entries = [re.sub(r'\d+|\([^)]+\)|\.', '', entry).strip() for entry in ex_cl_entries for sell in Seller if sell in entry]
        mr_entries = [re.sub(r'\d+|\([^)]+\)|\.', '', entry).strip() for entry in ex_cl_entries if '(MR)' in entry]
        dr_entries = [re.sub(r'\d+|\([^)]+\)|\.', '', entry).strip() for entry in ex_cl_entries if '(DR)' in entry]
        lr_entries = [re.sub(r'\d+|\([^)]+\)|\.', '', entry).strip() for entry in ex_cl_entries if '(LR)' in entry]
        rr_entries = [re.sub(r'\d+|\([^)]+\)|\.', '', entry).strip() for entry in ex_cl_entries if '(RR)' in entry]
        pl_entries = [re.sub(r'\d+|\([^)]+\)|\.', '', entry).strip() for entry in ex_cl_entries if '(PL)' in entry]
        me_entries = [re.sub(r'\d+|\([^)]+\)|\.', '', entry).strip() for entry in ex_cl_entries if '(ME)' in entry]
        de_entries = [re.sub(r'\d+|\([^)]+\)|\.', '', entry).strip() for entry in ex_cl_entries if '(DE)' in entry]
        le_entries = [re.sub(r'\d+|\([^)]+\)|\.', '', entry).strip() for entry in ex_cl_entries if '(LE)' in entry]
        re_entries = [re.sub(r'\d+|\([^)]+\)|\.', '', entry).strip() for entry in ex_cl_entries if '(RE)' in entry]
        ay_entries = [re.sub(r'\d+|\([^)]+\)|\.', '', entry).strip() for entry in ex_cl_entries if '(AY)' in entry]

        return {
            "SellerName": ', '.join(ex_entries),
            "PurchaserName": ', '.join(cl_entries),
            "Mortgage Requester": ', '.join(mr_entries),
            "Mortgage Executioner": ', '.join(me_entries),
            "Deed Requester": ', '.join(dr_entries),
            "Deed Executioner": ', '.join(de_entries),
            "Rent Requester": ', '.join(lr_entries),
            "Rent Executioner": ', '.join(le_entries),
            "Release Requester": ', '.join(rr_entries),
            "Release Executioner": ', '.join(re_entries),
            "Power of Attorney Giver": ', '.join(pl_entries),
            "Power of Attorney Accepter": ', '.join(ay_entries)
        }
    except Exception as e:
        print(f"Error extracting EX and CL: {str(e)}")
        return {
            "SellerName": "",
            "PurchaserName": "",
            "Mortgage Requester": "",
            "Mortgage Executioner": "",
            "Deed Requester": "",
            "Deed Executioner": "",
            "Rent Requester": "",
            "Rent Executioner": "",
            "Release Requester": "",
            "Release Executioner": "",
            "Power of Attorney Giver": "",
            "Power of Attorney Accepter": ""
        }

# Add empty keys for contact details of seller and purchaser
for column in entity_columns:
    transaction_data[column] = ""

# Apply the function to create separate columns for EX and CL
ex_cl_data = transaction_data["Name of Parties Executant(EX) & Claimants(CL)"].apply(extract_ex_cl)
try:
    transaction_data["SellerName"] = ex_cl_data.apply(lambda x: x["SellerName"])
    transaction_data["PurchaserName"] = ex_cl_data.apply(lambda x: x["PurchaserName"])

    # transaction_data["Mortgage Requester"] = ex_cl_data.apply(lambda x: x["Mortgage Requester"])
    # transaction_data["Mortgage Executioner"] = ex_cl_data.apply(lambda x: x["Mortgage Executioner"])

    # transaction_data["Deed Requester"] = ex_cl_data.apply(lambda x: x["Deed Requester"])
    # transaction_data["Deed Executioner"] = ex_cl_data.apply(lambda x: x["Deed Executioner"])

    # transaction_data["Rent Requester"] = ex_cl_data.apply(lambda x: x["Rent Requester"])
    # transaction_data["Rent Executioner"] = ex_cl_data.apply(lambda x: x["Rent Executioner"])

    # transaction_data["Release Requester"] = ex_cl_data.apply(lambda x: x["Release Requester"])
    # transaction_data["Release Executioner"] = ex_cl_data.apply(lambda x: x["Release Executioner"])

    # transaction_data["Power of Attorney Giver"] = ex_cl_data.apply(lambda x: x["Power of Attorney Giver"])
    # transaction_data["Power of Attorney Accepter"] = ex_cl_data.apply(lambda x: x["Power of Attorney Accepter"])

except KeyError as e:
    print(f"KeyError: {e}")

# Function to tag property type
def tag_property_type(description):
    try:
        for keyword in Commercial:
            if keyword in description:
                return "commercial"
        return "residential"
    except Exception as e:
        print(f"Error tagging property type: {str(e)}")
        return "residential"

# Define a function to extract document type
def extract_document_type(text):
    try:
        match = re.search(document_type_pattern, text)
        if match:
            document_type = match.group(2).strip()
            return document_type
        else:
            return None
    except Exception as e:
        print(f"Error extracting document type: {str(e)}")
        return None

# Define a function to extract value
def extract_value(text, value_type):
    try:
        value_pattern = r'{}:Rs.\s+([\d,]+)'.format(value_type)
        match = re.search(value_pattern, text)
        return match.group(1).replace(',', '') if match else None
    except Exception as e:
        print(f"Error extracting value for {value_type}: {str(e)}")
        return None

# Define a function to extract property details
def extract_details(text):
    try:
        locality_match = re.search(locality_pattern, text)
        house_match = re.search(house_pattern, text)
        apartment_match = re.search(apartment_pattern, text)
        flat_match = re.search(flat_pattern, text)
        built_match = re.search(built_pattern, text)
        floor_match01 = re.search(floor_pattern, text)
        floor_match02 = re.search(floor_pattern02, text)

        house = house_match.group(1) if house_match else None
        apartment = apartment_match.group(1) if apartment_match else None
        flat = flat_match.group(1) if flat_match else None
        built = built_match.group(1) if built_match else None
        floor2 = floor_match02.group(1) if floor_match02 else None
        floor1 = floor_match01.group(1) if floor_match01 else None
        locality = locality_match.group(1) if locality_match else None

        if apartment:
            apartment_name = apartment
        else:
            apartment_name = "Building name not available"

        # Combine floor numbers
        floor_combined = f"{floor1} {floor2}" if floor1 and floor2 else floor1 or floor2
        # Concatenate fields to form address

        address = f"{apartment_name}, {flat}, {locality},{'SECUNDERABAD'} {'Hyderabad'}, {'Telangana'}"
        building_address = f"{apartment_name}, {locality}, {'Hyderabad'}, {'Telangana'}"

        # Add empty keys as required
        for bek in transaction_empty_keys:
            transaction_data[bek] = ""

        return {
            "locality": locality,
            "VillageName": village,
            "PlotNo": house,
            "BuildingName": apartment_name,
            "FlatNo": flat,
            "totalSquareFeet": built,
            "MalaNo": floor_combined,
            "Status": Status,
            "city": city,
            "state": state,
            "Name": apartment_name,
            "Address": address,
            "Building Address": building_address
        }

    except Exception as e:
        print(f"Error extracting property details: {str(e)}")
        return {
            "locality": "",
            "VillageName": "",
            "PlotNo": "",
            "BuildingName": "",
            "FlatNo": "",
            "totalSquareFeet": "",
            "MalaNo": "",
            "Status": "",
            "city": "",
            "state": "",
            "Name": "",
            "Address": ""
        }

# Rename columns
transaction_data.rename(columns={
    "Description of property": "Property Description",
    "Name of Parties Executant(EX) & Claimants(CL)": "Name of Seller and Buyer",
    "Document Number": "DocumentNo",
    "Reg.Date": "RegistrationDate",
    "Exe.Date": "DocumentSubmissionDate",
    "Pres.Date": "PresentationDate"
}, inplace=True)

transaction_data = transaction_data.copy()

# Apply the function to extract property details
transaction_data["Extracted Details"] = transaction_data["Property Description"].apply(extract_details)

# NAME AND TYPE OF DOCUMENT ARE SAME
transaction_data["DocumentName"] = transaction_data["Nature, Mkt.Value, Con. Value"].apply(extract_document_type)  # NAME OF DOCUMENT
transaction_data["DocumentType"] = transaction_data["Nature, Mkt.Value, Con. Value"].apply(extract_document_type)  # DOCUMENT TYPE

#TYPE OF BUILDING "RESIDENTIAL" & "COMMERCIAL"
transaction_data["buildingType"] = transaction_data["Property Description"].apply(tag_property_type)

#MARKET PRICE AND CONSTRUCTION VALUE
transaction_data["MarketPrice"] = transaction_data["Nature, Mkt.Value, Con. Value"].apply(extract_value, args=("Mkt.Value",))
transaction_data["Construction Value"] = transaction_data["Nature, Mkt.Value, Con. Value"].apply(extract_value, args=("Cons.Value",))
transaction_data = pd.concat([transaction_data, transaction_data["Extracted Details"].apply(pd.Series)], axis=1)

# DROPPING EXTRACTED DETAILS AS IT IS NOT REQUIRED
transaction_data = transaction_data.drop(columns=["Extracted Details"])

# REMOVING ANY EXTRA WHITE SPACES IN THE DATAFRAME
transaction_data = transaction_data.applymap(strip_whitespace)

# DROPPING ALL THE DUPLICATES
transaction_data = transaction_data.drop_duplicates()

# creating a new dataframe for building data using transaction data
building_data = transaction_data[["BuildingName", "Property Description", "VillageName", "locality", "Building Address","PlotNo","totalSquareFeet","FlatNo","MalaNo","Status","buildingType","city","state","MarketPrice"]]
for column in building_empty_keys:
    building_data[column] = ""

#DROPPING BUILDING ADDRESS KEY
transaction_data = transaction_data.drop(columns=["Building Address"])

#RENAMING BUILDING ADDRESS KEY AS ADDRESS and MarketPrice key IN THE BUILDING DATA DATAFRAME
building_data.rename(columns={"Building Address": "Address",
                              "MarketPrice":"marketPrice"}, inplace=True)  # Rename the column

#saving the TRANSACTION DATA
try:
    transaction_data.to_excel('SECUNDERABAD_Transaction_Data.xlsx', index=False)
    print("Data saved successfully.")
except Exception as e:
    print(f"Error saving data: {str(e)}")

#SAVING THE BUILDING DATA
try:
    building_data.to_excel('SECUNDERABAD_BUILDING_DATA.xlsx', index=False)
    print("Data saved successfully.")
except Exception as e:
    print(f"Error saving data: {str(e)}")