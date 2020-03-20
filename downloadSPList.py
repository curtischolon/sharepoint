from shareplum import Site
from requests_ntlm import HttpNtlmAuth
from credentials import Credentials
import pandas as pd
import os


def download_dealer_list():
    ''' downloads the current dealer listing from SharePoint and saves as CustomerDocuments.xlsx '''
    creds = Credentials()

    cred = HttpNtlmAuth(creds.username, creds.password)
    site = Site('http://tng/SharedServices/AR', auth=cred)

    rows = site.List('Document_Circulation').get_list_items()

    columns = [
        'Service_Location',
        'Dealer_Name',
        'Master',
        'Chain_Master',
        'Dealer_Number',
        'Statements',
        'Invoices',
        'Credits',
        'Fax Number',
        'Attn To:',
        'Output',
        'Created',
        'Modified',
        'Item Type',
        'Path',
        'Email Address'
    ]

    dealer_list = []
    keys = []
    for row in rows:
        keys = keys + list(row.keys())
        # keys = list(row.keys())
        # r = [row[i] for i in keys]
        r = [
            row.get('Service_Location'),
            row.get('Dealer_Name'),
            row.get('Master'),
            row.get('Chain_Master'),
            row.get('Dealer_Number'),
            row.get('Statements'),
            row.get('Invoices'),
            row.get('Credits'),
            row.get('Fax Number', ''),
            row.get('Attn To:', ''),
            row.get('Output', ''),
            row.get('Created'),
            row.get('Modified'),
            row.get('Item Type'),
            row.get('Path', ''),
            row.get('Email Address', '')
        ]
        dealer_list.append(r)

    keys = set(keys)
    print(keys)
    df = pd.DataFrame(dealer_list, columns=columns)
    # delete existing listing, if it exists
    if os.path.isfile('../Customer_Documents.xlsx'): os.remove('../Customer_Documents.xlsx')
    df.to_excel('../Customer_Documents.xlsx', index=False, header=True)

def main():
    download_dealer_list()

if __name__ == "__main__":
    main()




