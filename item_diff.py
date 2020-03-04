# -*- coding: utf-8 -*-
"""
Created on Mon Dec  2 10:53:45 2019

@author: torong
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Oct 31 15:23:28 2019

@author: torong
"""

import sys
import os
import pandas as pd
import datetime as dt
import_dir = (os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
+ '/Support/')
sys.path.append(import_dir)
from commonOperations import ImportOperations
from ExcelSupport import ExcelSupport
#OEM_item_data = r"C:\Users\torong\Desktop\Item_status_diff\OEM_item_data.xlsx"

class ItemMasterDataReconcilliation:
    def __init__(self):
        self.item_master_data = {'SC01': r"C:\Users\torong\Desktop\SCRIPT\MASTER_DATA\Masters - Items - SC01 - EMEA.xlsx", 
                            'SC02': r"C:\Users\torong\Desktop\SCRIPT\MASTER_DATA\Masters - Items - SC02 - EMEA.xlsx", 
                            'SC03': r"C:\Users\torong\Desktop\SCRIPT\MASTER_DATA\Masters - Items - SC03 - EMEA.xlsx", 
                            'SC04': r"C:\Users\torong\Desktop\SCRIPT\MASTER_DATA\Masters - Items - SC04 - EMEA.xlsx", 
                            'SC05': r"C:\Users\torong\Desktop\SCRIPT\MASTER_DATA\Masters - Items - SC05 - EMEA.xlsx"}
        self.headers = ['ItemId', 'ProductName', 'ProductId', 'ItemResponsible',
                        'CostingSupplier', 'Sales_Stopped', 'AltItemId',
                        'SearchName', 'ItemReqGroup',
                        'Purch_Stopped', 'Purch_Price1', 'Purch_PriceUnit1',
                        'Sales_Price', 'Sales_PriceUnit']
        
        
    def import_data(self, domain1, domain2, update_database = False):
        
        domain_path1 = self.item_master_data[domain1]
        domain_path2 = self.item_master_data[domain2]
        
        if update_database:
            io = ImportOperations()
            print(domain_path1)
            print('Updating database 1. Do not interfere')
            io.update_database(domain_path1)
            print('Updating database 2. Do not interfere')
            io.update_database(domain_path2)
        
        print('Reading data')
        df_domain1 = pd.read_excel(domain_path1, header=1, usecols=self.headers)
        df_domain2 = pd.read_excel(domain_path2, header=1, usecols=self.headers)
        return(df_domain1, df_domain2)
        
    def find_master_data_errors(self, df1, df2):
        #df1 is the is the dependent domain while df2 is the dominant domain
        #That means that information is sent from df1, while df2 is the party that interprets the data
        
        print(len(df1),'1 <- Lines proecessed in Domain -> 2 ', len(df2))
        productID_with_error = {}
        
        for i in range(len(df1['ProductId'])):
            item_error = {
                    
                    'Item ID Domain1': None,
                    'Item ID Domain2': None,
                    'Item search name': None,
                    'Item Req Group': None,
                    'Alt Item ID': None,
                    'Multiple item ID refer to product ID': None,
                    'Sales stopped, purch open': None,
                    'Item ID not identical': None,
                    'Price mismatch': None,
                    'Price unit quantity mismatch': None,
                    'Product missing in Domain 2': None,
                    
                    }
            pair_index = None
            for j in range(len(df2['ProductId'])):
                if df1['ProductId'][i] == df2['ProductId'][j]:
                    D1, D2 = df1.iloc[i], df2.iloc[j]
                    
                    if D1['ProductId'] in productID_with_error.keys():
                        productID_with_error['ProductId']['Multiple item ID refer to product iD'] = D2['ItemId']
                    
                    else:
                        if D1['Sales_Stopped'] != D2['Purch_Stopped']:
                            item_error['Sales stopped, purch open'] = True
                        if str(D1['ItemId']) != str(D2['ItemId']):
                           item_error['Item ID not identical'] = (str(D1['ItemId']), str(D2['ItemId']))
                        if D1['Purch_Price1'] != D2['Sales_Price']:
                            item_error['Price mismatch'] = (D1['Sales_Price'], D2['Sales_Price'])
                        if D1['Purch_PriceUnit1'] != D2['Sales_PriceUnit']:
                            item_error['Price unit quantity mismatch'] = (D1['Purch_PriceUnit1'], D2['Sales_PriceUnit'])
                        pair_index = (i,j)
                    
                        
            if not pair_index:
                item_error['Product missing in Domain 2'] = True
            
            if item_error:
                item_error['Item search name'] = df1['SearchName'].iloc[i]
                item_error['Alt Item ID'] = df1['AltItemId'].iloc[i]
                item_error['Item ID Domain1'] = df1['ItemId'].iloc[i]
                item_error['Item responsible1'] = df1['ItemResponsible'].iloc[i]
                item_error['Item Req Group'] = df1['ItemReqGroup'].iloc[i]
                productID_with_error[df1['ProductId'].iloc[i]] = item_error
                if not item_error['Product missing in Domain 2']:
                    item_error['Item ID Domain2'] = df2['ItemId'].iloc[pair_index[1]]
                    item_error['Item responsible2'] = df2['ItemResponsible'].iloc[pair_index[1]]

            
        return(productID_with_error)
                
                            
                           
                            

if __name__ == '__main__':
    imdr = ItemMasterDataReconcilliation()
    a, b = imdr.import_data("SC02", "SC03", update_database=False)
    
    a0 = a.loc[a['CostingSupplier'] == 200150].reset_index()
    a0 = a0.rename(columns = {'index': "Org indx"})
    b0 = b.rename(columns = {'index': "Org indx"})
    
 
    error = imdr.find_master_data_errors(a0,b0)
    
    wrtr = ExcelSupport
    wrtr.write_a_nested_dict_to_excel(error)
    
