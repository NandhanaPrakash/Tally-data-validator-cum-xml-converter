import pandas as pd
import xml.etree.ElementTree as ET
import uuid
from datetime import datetime

# Loading xlsx file and reading all sheets
xlsx_file_path = "C:\\Users\\Dell\\Downloads\\voucher_data.xlsx"
sheets = pd.read_excel(xlsx_file_path, sheet_name=None)

# xlsx to csv converter
def xlsx_to_csv(df, csv_file):
    df.to_csv(csv_file, index=False)
    print(f'XML Conversion unsuccessful due to invalid data. CSV file saved at: {csv_file}')

# xlsx to xml converter
def xlsx_to_xml(df, output_file_path):
    df = df.fillna({'Reference No': '', 'Narration': '', 'Ledger Name': '', 'Ledger Name.1': '', 'Amount': 0, 'Amount.1': 0})

    envelope = ET.Element('ENVELOPE')

    header = ET.SubElement(envelope, 'HEADER')
    tallyrequest = ET.SubElement(header, 'TALLYREQUEST')
    tallyrequest.text = 'Import Data'

    body = ET.SubElement(envelope, 'BODY')
    importdata = ET.SubElement(body, 'IMPORTDATA')

    requestdesc = ET.SubElement(importdata, 'REQUESTDESC')
    reportname = ET.SubElement(requestdesc, 'REPORTNAME')
    reportname.text = 'Vouchers'
    staticvariables = ET.SubElement(requestdesc, 'STATICVARIABLES')
    svccompany = ET.SubElement(staticvariables, 'SVCURRENTCOMPANY')
    svccompany.text = 'Raj'

    requestdata = ET.SubElement(importdata, 'REQUESTDATA')

    for index, row in df.iterrows():
        tallymessage = ET.SubElement(requestdata, 'TALLYMESSAGE', {'xmlns:UDF': 'TallyUDF'})
        voucher = ET.SubElement(tallymessage, 'VOUCHER', {
            'REMOTEID': f'aaeb8870-afa4-4fe9-bd6f-0f75021f37b5-3RAJ14500-{index}',
            'VCHTYPE': row['Voucher Type'],
            'ACTION': 'Create'
        })

        vouchertypename = ET.SubElement(voucher, 'VOUCHERTYPENAME')
        vouchertypename.text = row['Voucher Type']

        date_obj = pd.to_datetime(row['Date'], dayfirst=True)
        date_str = date_obj.strftime('%Y%m%d')

        date = ET.SubElement(voucher, 'DATE')
        date.text = date_str

        effectivedate = ET.SubElement(voucher, 'EFFECTIVEDATE')
        effectivedate.text = date_str

        reference = ET.SubElement(voucher, 'REFERENCE')
        reference.text = str(row['Reference No'])

        narration = ET.SubElement(voucher, 'NARRATION')
        narration.text = row['Narration']

        guid = ET.SubElement(voucher, 'GUID')
        guid.text = f'aaeb8870-afa4-4fe9-bd6f-0f75021f37b5-3RAJ14500-{index}'

        alterid = ET.SubElement(voucher, 'ALTERID')
        alterid.text = str(14500 + index)

        for ledger, amount, effect in zip(
            [row['Ledger Name'], row['Ledger Name.1']],
            [-row['Amount'], row['Amount.1']],
            [row['Effect'], row['Effect.1']]
        ):
            if pd.isna(ledger):
                continue
            allledgerentries = ET.SubElement(voucher, 'ALLLEDGERENTRIES.LIST')

            removezeroentries = ET.SubElement(allledgerentries, 'REMOVEZEROENTRIES')
            removezeroentries.text = 'No'

            isdeemedpositive = ET.SubElement(allledgerentries, 'ISDEEMEDPOSITIVE')
            isdeemedpositive.text = 'Yes' if effect == 'Dr.' else 'No'

            ledgername = ET.SubElement(allledgerentries, 'LEDGERNAME')
            ledgername.text = ledger

            amount_el = ET.SubElement(allledgerentries, 'AMOUNT')
            amount_el.text = str(amount)

    xml_str = ET.tostring(envelope, encoding='unicode')

    with open(output_file_path, 'w') as f:
        f.write(xml_str)

    print(f"XML file has been created: {output_file_path}")

# Function returning the number of days in each month
def days_in_month(year, month):
    if month in [1, 3, 5, 7, 8, 10, 12]:
        return 31
    elif month in [4, 6, 9, 11]:
        return 30
    elif month == 2:
        # Checking for leap year
        if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0):
            return 29
        else:
            return 28
    else:
        return 0

# List defining allowed voucher values
allowed_voucher_values = {"Journal", "Contra", "Receipt", "Payment"}

# Iterate over each sheet
for sheet_name, df in sheets.items():
    print(f"Processing sheet: {sheet_name}")

    # Converting date to parse-able format
    df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], format="%d-%m-%y", dayfirst=True, errors='coerce')

    if df.iloc[:, 0].isnull().any():
        print("Some dates could not be parsed and have been set to NaT.")
        #print(df[df.iloc[:, 0].isnull()])

    # Stores rows with invalid voucher type after checking
    invalid_voucher_cells = []
    for index, value in df.iloc[:, 1].items():  # Index 1 = column 2
        if value not in allowed_voucher_values:
            invalid_voucher_cells.append((index + 2, value))

    # Checking the month column for invalid values (1-12 only allowed)
    invalid_month_cells = []
    for index, value in df.iloc[:, 3].items():  # Index 3 = 4th column
        if not (1 <= value <= 12):
            invalid_month_cells.append((index + 2, value))

    # Stores rows with invalid date value after checking
    invalid_date_cells = []
    for index, row in df.iterrows():
        try:
            day = int(row.iloc[2])  # Date value in the third column
            month = int(row.iloc[3])  # Month value in the fourth column
            year = row.iloc[0].year  # Year from the first column

            if not (1 <= day <= days_in_month(year, month)):
                invalid_date_cells.append((index + 2, day, month))

        except ValueError:
            continue

    # Checking the Amount columns (dr and cr effect) for matching values and positive numbers
    invalid_amount_cells = []
    for index, row in df.iterrows():
        amount1 = row.iloc[7]  # First amount column H
        amount2 = row.iloc[10]  # Second amount column K
        if amount1 != amount2 or amount1 <= 0 or amount2 <= 0:
            invalid_amount_cells.append((index + 2, amount1, amount2))

    # Checking the Reference No column for values starting with 1 and incremented by 1 for every consecutive row
    invalid_reference_cells = []
    expected_reference_number = 1
    for index, value in df.iloc[:, 4].items():
        if value != expected_reference_number:
            invalid_reference_cells.append((index + 2, value))
        expected_reference_number += 1

    # Checking the effect columns for Receipt and Contra voucher types
    invalid_effect_cells = []
    for index, row in df.iterrows():
        voucher_type = row.iloc[1]  # Voucher Type column B
        effect1 = row.iloc[6]  # First effect column G
        effect2 = row.iloc[9]  # Second effect column J
        if voucher_type in {"Receipt", "Contra"}:
            if effect1 != "Cr." or effect2 != "Dr.":
                invalid_effect_cells.append((index + 2, voucher_type, effect1, effect2))
        else:
            if not ((effect1 == "Cr." and effect2 == "Dr.") or (effect1 == "Dr." and effect2 == "Cr.")):
                invalid_effect_cells.append((index + 2, voucher_type, effect1, effect2))
        if voucher_type == "Journal":
            if effect1 != "Dr." or effect2 != "Cr.":
                invalid_effect_cells.append((index + 2, voucher_type, effect1, effect2))

    # Check for validation errors and save to CSV if any
    if invalid_voucher_cells or invalid_month_cells or invalid_date_cells or invalid_amount_cells or invalid_reference_cells or invalid_effect_cells:
        csv_file_path = f"C:\\Users\\Dell\\Downloads\\{sheet_name}_voucher_data.csv"
        xlsx_to_csv(df, csv_file_path)

        if invalid_voucher_cells:
            print("Invalid 'Voucher Type' data found in the following cells:")
            for row, value in invalid_voucher_cells:
                print(f"Row {row}: {value}")

        if invalid_month_cells:
            print("\nInvalid 'Month' data found in the following cells:")
            for row, value in invalid_month_cells:
                print(f"Row {row}: {value}")

        if invalid_date_cells:
            print("\nInvalid 'Date' data found in the following cells:")
            for row, day, month in invalid_date_cells:
                print(f"Row {row}: {day}/{month}")

        if invalid_amount_cells:
            print("\nInvalid 'Amount' data found in the following cells:")
            for row, amount1, amount2 in invalid_amount_cells:
                print(f"Row {row}: {amount1}, {amount2}")

        if invalid_reference_cells:
            print("\nInvalid 'Reference No' data found in the following cells:")
            for row, value in invalid_reference_cells:
                print(f"Row {row}: {value} (Expected {row - 1})")

        if invalid_effect_cells:
            print("\nInvalid 'Effect' data found in the following cells:")
            for row, voucher_type, effect1, effect2 in invalid_effect_cells:
                print(f"Row {row}: {voucher_type} - Effect1: {effect1}, Effect2: {effect2}")

    else:
        print("All data is valid.")
        xml_file_path = f"C:\\Users\\Dell\\Downloads\\{sheet_name}_voucher_data.xml"
        xlsx_to_xml(df, xml_file_path)
        print(f"Validated xml file found at: {xml_file_path}")
