# -*- coding: utf-8 -*-
"""
Created on Mon Feb  3 16:24:11 2020

@author: torong
"""
import pandas as pd
import datetime as dt
import os
import sys

import_dir = (os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
+ '/Support/')
sys.path.append(import_dir)
import ImportOperations as IO

def add_master_data_to_parent(parent_item_dictionary, master_path ,update_master=False):
    
    md = IO.MasterDataOperations
    if update_master:
        md.update_masterdata(master_path)
    
    md_df = md.import_master_data_by_columns(path=master_path, column_headers_or_numbers=['ProductId', 'ItemResponsible', 'Sales_Stopped'])
        
    for parent in parent_item_dictionary.keys():
        try:
            md_df.loc['ProductId']
        except:
            pass



if __name__ == '__main__':
    #BOM_path_01 = r"C:\Users\torong\Desktop\SCRIPT\BOM structure\Sc01 BOM struktur.xlsx"
    #BOM_path_02 = r"C:\Users\torong\Desktop\SCRIPT\BOM structure\Sc02 BOM struktur.xlsx"
    #BOM_path_03 = r"C:\Users\torong\Desktop\SCRIPT\BOM structure\Sc03 BOM struktur.xlsx"
    #BOM_path_05 = r"C:\Users\torong\Desktop\SCRIPT\BOM structure\Sc05 BOM struktur.xlsx"
    #BOM_path_test = r"C:\Users\torong\Desktop\SCRIPT\BOM structure\BOM_test.xlsx"
    test_domain_paths = [r"C:\Users\torong\Desktop\SCRIPT\BOM structure\Sc02 BOM struktur test.xlsx"]
    domain_paths = [r"C:\Users\torong\Desktop\SCRIPT\BOM structure\Sc02 BOM struktur.xlsx"]
    #domain_paths = test_domain_paths
    
    info_columns = ['ParentBomId', 'ParentItemId', 'ParentName', 'ParentProductNumber',
                   'ParentAltItemID', 'ChildAltItemID','ChildItemId','ChildName','ChildBomId',
                   'ChildProductNumber', 'ParentSalesWareHouse', 'ChildBomWareHouse',  'ChildSalesWareHouse',
                   'ChildCostingSupplier', 'ProdFlushingPrinciple', 'Qty']
   
    error_list=[]
    df_list = []
    time = str(dt.datetime.now())[:16].replace('-', '_').replace(':', '_').replace(' ', '_')
    writer = pd.ExcelWriter('BOM_{}.xlsx'.format(time))
    for i in range(len(domain_paths)):
        BOM_path = domain_paths[i]
        sheet = 'BOM'
        
        rd = IO.ImportBOM
        update = False
        if update:
            rd.update_BOM_data(BOM_path)
        
        df = rd.import_BOM_by_columns(path = BOM_path, sheet_name=sheet,
                                              column_headers_or_numbers=info_columns)
        df = df.fillna(value=0)
        df['ParentSalesWareHouse'] = df['ParentSalesWareHouse'].astype('int64').astype(str)
        df['ChildCostingSupplier'] = df['ChildCostingSupplier'].astype('int64').astype(str)
        BOM_prod_dict = {}
        for i, BOMid in enumerate(df['ParentBomId'].unique()):
            BOM_prod_dict[BOMid] = {
                    'DISC Parent': False,
                    'Parent BOM': None,
                    'Active BOM': True,
                    'Parent': None,
                    'BOM_Warehouse': None,
                    'BOM_FlushingPrinciple': None,
                    'External Parent info': None,
                    'Flushing principle': None,
                    
                    'Error boolean' : False,
                    'FP Error': False,
                    'WH error': False,
                    'Child duplicates': False,
                    'Is child': False,
                    
                    
                    
                    'Children' : [],
                    'Error' : [],
                    }
        
        for b, bom in enumerate(df['ParentBomId']):
            if df['ParentItemId'][b] == 0:
               BOM_prod_dict[bom]['Active BOM'] == False 
           
                
            
            parentID = str(df['ParentItemId'][b])
            parentPID = str(df['ParentProductNumber'][b])
            parentName = str(df['ParentName'][b])
            parentAltItem = str(df['ParentAltItemID'][b])
            parentBOMwh = str(df['ParentSalesWareHouse'][b])
            childID = str(df['ChildItemId'][b])
            childPID = str(df['ChildProductNumber'][b])
            childName = str(df['ChildName'][b])
            childAltItem = str(df['ChildAltItemID'][b])
            childBOMID =str(df['ChildBomId'][b])
            childCostingSupplier = str(df['ChildCostingSupplier'][b])
            childSalesWH = str(df['ChildSalesWareHouse'][b])
            childBOMWH = str(df['ChildBomWareHouse'][b])
            childFlushinPrinciple = str(df['ProdFlushingPrinciple'][b])
            qty = str(df['Qty'][b])
            
            BOM_error_info_list = []
            
            
            
            #If the current line does not have a parent ID, it is the assigned the parent ID of the previous 
            DISC_labels = ['Nothing in stock', 'Always']
            if not BOM_prod_dict[bom]['Parent']:
                BOM_prod_dict[bom]['Parent'] = parentID
                #Add a flag for DISC items
                for label in DISC_labels:
                    if label in parentAltItem:
                        BOM_prod_dict[bom]['DISC Parent'] = True
                #Add a flag for parent beeing a child
                if parentID in df['ChildItemId']:
                    BOM_prod_dict[bom]['Is child'] = True
                    
    
        
            #Check for a flushing principle and transform from Numeric to terminology
            if childFlushinPrinciple == '0':
                BOM_error_info_list.append('Child {} is missing Flusing principle'.format(childID))
                FlushingPrinciple = 'Missing'
                BOM_prod_dict[bom]['FP Error'] = True
                
            elif childFlushinPrinciple == '3':
                FlushingPrinciple = 'Finished'
            elif childFlushinPrinciple == '1':
                FlushingPrinciple = 'Manual'
            else:
                FlushingPrinciple = 'Error'
            
            #Test and check that there isn't more than one child flushinprinciple
            if not BOM_prod_dict[bom]['BOM_FlushingPrinciple']:
                BOM_prod_dict[bom]['BOM_FlushingPrinciple'] = childFlushinPrinciple
            else:
                if BOM_prod_dict[bom]['BOM_FlushingPrinciple'] != childFlushinPrinciple:
                    BOM_error_info_list.append('FP is not uniform. Child {0} FP {1} does not match the first FP registered for this BOM {2}'.format(
                            childID, childFlushinPrinciple, BOM_prod_dict[bom]['BOM_FlushingPrinciple']))
                    BOM_prod_dict[bom]['FP Error'] = True 
    
    
            #A flag for BOM that refers to multiple parent products
            if parentID != BOM_prod_dict[bom]['Parent']:
                BOM_error_info_list.append('BOM ID refer to multiple products')
                
            #If the child item has already been added to the BOM it indicates duplicates.
            #Checks for multiple child items in BOM
            if childID in BOM_prod_dict[bom]['Children']:
                BOM_error_info_list.append('Multiple entries of the same child in BOM')
                BOM_prod_dict[bom]['Child duplicates'] = True 
               
            #Check if the child item in the BOM has been discontinued
            #If yes, append error message and child, else append child
            if  childAltItem in ['Always','Nothing in stock']:
                BOM_prod_dict[bom]['Children'].append(childID)
                BOM_error_info_list.append('The Child has been Discontinued {}'.format(childID))
            else:
                BOM_prod_dict[bom]['Children'].append(childID)
            
            #Check if the child has a WH
            if not childBOMID:
                BOM_error_info_list.append('The Child {} WH is missing'.format(childID))
                BOM_prod_dict[bom]['WH error'] = True
            #Check if the child WH is consistent/uniform on all lines
            if not BOM_prod_dict[bom]['BOM_Warehouse']:
                BOM_prod_dict[bom]['BOM_Warehouse'] = childBOMWH
            if BOM_prod_dict[bom]['BOM_Warehouse'] != childBOMWH:
                BOM_error_info_list.append('The Child {} WH does not match the first registered child WH registered: {}'.format(childBOMWH, BOM_prod_dict[bom]['BOM_Warehouse'] ))
                BOM_prod_dict[bom]['WH error'] = True
                
            #Info pulled straight from the BOM info
            BOM_info = [
                    bom, parentID, parentName, parentPID,
                    parentAltItem, childAltItem, childID, childName, childBOMID,
                    childPID, parentBOMwh, childSalesWH,  childBOMWH, childCostingSupplier, childFlushinPrinciple, qty
                    ]
            
            #Status indicators, processed from the BOM info
            processed_columns=['DISC_PARENT' , 'BOM_ACTIVE', 'Flushing principle status', 'FP err', 'Is child', 'WH err', 'Child dup' ]
            processed_info = [BOM_prod_dict[bom]['DISC Parent'],
                              BOM_prod_dict[bom]['Active BOM'], FlushingPrinciple, 
                              BOM_prod_dict[bom]['FP Error'], BOM_prod_dict[bom]['Is child'],
                              BOM_prod_dict[bom]['WH error'], BOM_prod_dict[bom]['Child duplicates']]
            
            if len(BOM_error_info_list)>0:    
                error_info_message = BOM_info + processed_info + BOM_error_info_list
                BOM_prod_dict[bom]['Error'].append(BOM_error_info_list)
                error_list.append(error_info_message)
    
        error_pd = pd.DataFrame(error_list)
        error_pd_head = error_pd.head()
        w = 0
        header = info_columns + processed_columns
        while len(error_pd.columns)>len(header):
            w+=1
            header.append('Error message {}'.format(w))
        
        #Dataframe created from the dictionary 
        error_pd = pd.DataFrame(error_list)
        error_pd.columns = header
        df_list.append(error_pd)
        error_pd.to_excel(writer, sheet)
        
    writer.save()
