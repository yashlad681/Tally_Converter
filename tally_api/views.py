import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime
from django.core.files.storage import default_storage
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status


class ExportToXl(APIView):
    def convert_to_excel(self,file):
        tree = ET.parse(file)
        root = tree.getroot()
        body = root.findall('BODY')[0]

        # Create a Pandas dataframe from some data
        column_names = ['Date', 'Transaction Type', 'Vch No.', 'Ref No', 'Ref Type', 'Ref Date', 'Debtor', 'Ref Amount', 'Amount', 'Particulars', 'Vch Type', 'Amount Verified']

        # Convert the list of strings to a Pandas dataframe
        df = pd.DataFrame(columns=column_names)

        # Write the dataframe to an Excel file
        df.to_excel('result.xlsx', index=False)

        for child in body.findall('IMPORTDATA')[0].findall('REQUESTDATA')[0].findall('TALLYMESSAGE'):
            pairs = child.find('VOUCHER')
            if pairs.get('VCHTYPE')=="Receipt":
                date = datetime.strptime( pairs.find('EFFECTIVEDATE').text, "%Y%m%d")
                formatted_date = date.strftime("%d-%m-%Y")
                parent = {'Date': formatted_date, 'Transaction Type': 'Parent', 'Vch No.': pairs.find('VOUCHERNUMBER').text, 'Ref No': 'NA', 
                'Ref Type': 'NA', 'Ref Date': 'NA', 'Debtor': pairs.find('PARTYLEDGERNAME').text, 'Ref Amount': 'NA', 'Amount': None,
                'Particulars': pairs.find('PARTYLEDGERNAME').text, 'Vch Type': pairs.find('VOUCHERTYPENAME').text, 'Amount Verified': 'No'}
                total_amount = 0
                for entries in pairs.find('ALLLEDGERENTRIES.LIST').findall('BILLALLOCATIONS.LIST'):
                    total_amount += float(entries.find('AMOUNT').text)

                for entries in pairs.findall('ALLLEDGERENTRIES.LIST'):
                    if entries.find('ISDEEMEDPOSITIVE').text == 'No':
                        parent['Amount'] = entries.find('AMOUNT').text
                        if total_amount == float(parent['Amount']):
                            parent['Amount Verified'] = 'Yes'
                        df = df.append(parent, ignore_index=True)
                        df.to_excel('result.xlsx', index=False)
                        for child in entries.findall('BILLALLOCATIONS.LIST'):
                            df = df.append( {'Date': formatted_date, 'Transaction Type': 'Child', 
                            'Vch No.': pairs.find('VOUCHERNUMBER').text,
                            'Ref No': child.find('NAME').text, 'Ref Type': child.find('BILLTYPE').text, 
                            'Ref Date': '', 'Debtor': pairs.find('PARTYLEDGERNAME').text, 
                            'Ref Amount': child.find('AMOUNT').text, 'Amount': 'NA',
                            'Particulars': pairs.find('PARTYLEDGERNAME').text, 
                            'Vch Type': pairs.find('VOUCHERTYPENAME').text, 'Amount Verified': 'NA'}, ignore_index=True)
                            df.to_excel('result.xlsx', index=False)
                    else:
                        df = df.append( {'Date': formatted_date, 'Transaction Type': 'Other', 
                        'Vch No.': pairs.find('VOUCHERNUMBER').text, 'Ref No': 'NA', 
                            'Ref Type': 'NA', 'Ref Date': 'NA', 'Debtor': entries.find('LEDGERNAME').text, 
                            'Ref Amount': 'NA', 'Amount':  entries.find('AMOUNT').text,
                            'Particulars': entries.find('LEDGERNAME').text, 
                            'Vch Type': pairs.find('VOUCHERTYPENAME').text, 'Amount Verified': 'NA'}, ignore_index=True)
                        df.to_excel('result.xlsx', index=False)

        file_url = default_storage.url('result.xlsx')
        return file_url
        
    def post(self, request):
        try:
            file = request.FILES['file']
            file_url = self.convert_to_excel(file)
            json_obj = {"url": file_url}
            return Response(json_obj, status=status.HTTP_200_OK)
        except Exception as e:
            return Response({'debug_msg':str(e)}, status=status.HTTP_400_BAD_REQUEST)



