'''
Original Author:Nandhana Prakash

Date:22nd July 2024

This is to convert an excel file to xml file which will be loaded to tally

This function is called only when all validations are passed and the file is fit to be uploaded to Tally

The code is not perfect and may contain errors so feel free to provide feedback
'''

import pandas as pd
import uuid
from datetime import datetime

file_path = "C:\\Users\\Dell\\Downloads\\trial.xlsx"
df = pd.read_excel(file_path)

#Function that loads it in xml file for each row
def create_voucher_xml(row):
    voucher_type = row['Voucher Type']
    
    #Converting the date to string and then parsing it
    if isinstance(row['Date'], pd.Timestamp):
        date_str = row['Date'].strftime('%d/%m/%Y')
    else:
        date_str = row['Date']
    
    date = datetime.strptime(date_str, '%d/%m/%Y').strftime('%Y%m%d')  # Convert to 'YYYYMMDD' format
    
    reference = row['Reference No']
    ledger1 = row['Ledger Name']
    effect1 = row['Effect']
    amount1 = row['Amount']
    ledger2 = row['Ledger Name.1']
    effect2 = row['Effect.1']
    amount2 = row['Amount.1']
    narration = row['Narration']
    
    is_deemed_positive1 = 'Yes' if effect1 == 'Dr' else 'No'
    is_deemed_positive2 = 'No' if effect2 == 'Dr' else 'Yes'
    amount1 = f"-{amount1}" if effect1 == 'Dr' else f"{amount1}"
    amount2 = f"{amount2}" if effect2 == 'Dr' else f"-{amount2}"
    
    guid = str(uuid.uuid4())
    
    return f"""
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
        <VOUCHER REMOTEID="{guid}" VCHTYPE="{voucher_type}" ACTION="Create">
            <VOUCHERTYPENAME>{voucher_type}</VOUCHERTYPENAME>
            <DATE>{date}</DATE>
            <EFFECTIVEDATE>{date}</EFFECTIVEDATE>
            <REFERENCE>{reference}</REFERENCE>
            <NARRATION>{narration}</NARRATION>
            <GUID>{guid}</GUID>
            <ALLLEDGERENTRIES.LIST>
                <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
                <ISDEEMEDPOSITIVE>{is_deemed_positive1}</ISDEEMEDPOSITIVE>
                <LEDGERNAME>{ledger1}</LEDGERNAME>
                <AMOUNT>{amount1}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
                <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
                <ISDEEMEDPOSITIVE>{is_deemed_positive2}</ISDEEMEDPOSITIVE>
                <LEDGERNAME>{ledger2}</LEDGERNAME>
                <AMOUNT>{amount2}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
        </VOUCHER>
    </TALLYMESSAGE>"""

# Generating XML for each row and combining
xml_content = "<ENVELOPE><HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER><BODY><IMPORTDATA><REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME><STATICVARIABLES><SVCURRENTCOMPANY>Company 1</SVCURRENTCOMPANY></STATICVARIABLES></REQUESTDESC><REQUESTDATA>"
for index, row in df.iterrows():
    xml_content += create_voucher_xml(row)
xml_content += "</REQUESTDATA></IMPORTDATA></BODY></ENVELOPE>"

#Writing to xml file
output_file_path = "C:\\Users\\Dell\\Downloads\\trial.xml"
with open(output_file_path, 'w') as file:
    file.write(xml_content)




