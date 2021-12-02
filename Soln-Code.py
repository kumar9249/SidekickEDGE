import pandas as pd
import numpy as np

filename = input("Enter the file name: ")

if filename is not None:

    filepath = r'C:\Users\Lenovo\Desktop\SidekickEDGE\ ' + filename + '.xlsx'
    
    try:
        df1 = pd.read_excel(filepath, sheet_name=1, parse_dates=[5, 6, 7])
        df2 = pd.read_excel(filepath, sheet_name=2, parse_dates=[4, 5, 6])
        df3 = pd.read_excel(filepath, sheet_name=3, parse_dates=[6, 7, 8])
    
    except Exception:
        print("File not found")
        
    else:
        df1.rename(columns={'SERVICE_type': 'Service Type',
                             'Invoice Status': 'INVOICESTATUS',
                             'BILLED_DATE': 'Billed Date',
                             'COLLECTED': 'Collected Date',
                             'Unique ID check': 'Unique ID Check'}, inplace=True)
        df2.rename(columns={'SERVICE Type': 'Service Type',
                             'BILLED_DATE': 'Billed Date',
                             'COLLECTED': 'Collected Date',
                             'Unique ID check': 'Unique ID Check'}, inplace=True)
        
        new_df1 = df1.sort_values(by=['Billed Date', 'ID', 'Service Name'], ascending=[False, True, True])
        new_df2 = df2.sort_values(by=['Billed Date', 'ID'], ascending=[False, True])
        new_df3 = df3.sort_values(by=['Billed Date', 'ID', 'Service Name'], ascending=[False, True, True])
        
        master = pd.concat([new_df1, new_df2, new_df3], axis=0, join='outer', ignore_index=True)
        
        duplicates = master[master.duplicated(subset=['Billed Date', 'Service Type'], keep='first')]
        dup = duplicates[duplicates.duplicated('ID', keep='first')]
        
        def fillD(row):
            if row['ID'] in list(dup['ID']):
                return 'D'
            else:
                return np.nan
        master['Unique ID Check'] = master.apply(lambda row: fillD(row), axis=1)
        master.drop(columns=['Completion date'], axis=1, inplace=True)
        
        new_df1['Billed Date'] = pd.to_datetime(new_df1['Billed Date']).dt.date
        new_df2['Billed Date'] = pd.to_datetime(new_df2['Billed Date']).dt.date
        new_df3['Billed Date'] = pd.to_datetime(new_df3['Billed Date']).dt.date
        
        Service_1 = new_df1.pivot_table(values='ID', index='Billed Date', aggfunc='nunique')
        Service_1.rename(columns={'ID': 'Service 1'}, inplace=True)
        Service1 = new_df1.pivot_table(values='ID', index='Billed Date', aggfunc='count')
        Service1.rename(columns={'ID': 'Service 1'}, inplace=True)
        
        Service_2 = new_df2.pivot_table(values='ID', index='Billed Date', aggfunc='nunique')
        Service_2.rename(columns={'ID': 'Service 2'}, inplace=True)
        Service2 = new_df2.pivot_table(values='ID', index='Billed Date', aggfunc='count')
        Service2.rename(columns={'ID': 'Service 2'}, inplace=True)
        
        Service_3 = new_df3.pivot_table(values='ID', index='Billed Date', aggfunc='nunique')
        Service_3.rename(columns={'ID': 'Service 3'}, inplace=True)
        Service3 = new_df3.pivot_table(values='ID', index='Billed Date', aggfunc='count')
        Service3.rename(columns={'ID': 'Service 3'}, inplace=True)
        
        new_dfx1 = pd.concat([Service_1, Service_2, Service_3], axis=1)
        new_dfx2 = pd.concat([Service1, Service2, Service3], axis=1)
    
        new_dfx1 = new_dfx1.pivot_table(values=['Service 1', 'Service 2', 'Service 3'], index='Billed Date', aggfunc='sum')
        new_dfx2 = new_dfx2.pivot_table(values=['Service 1', 'Service 2', 'Service 3'], index='Billed Date', aggfunc='sum')
        
        new_dfx = pd.concat([new_dfx1, new_dfx2], axis=1)
        
        new_dfy = new_dfx2.copy()
        new_dfy['Count of records (Service 1 and Service 2)'] = new_dfy['Service 1'] + new_dfy['Service 2']
        new_dfy['Count of records (Service 2 and Service 3)'] = new_dfy['Service 2'] + new_dfy['Service 3']
        new_dfy['Count of records (Service 1 and Service 3)'] = new_dfy['Service 1'] + new_dfy['Service 3']
        new_dfy.drop(columns=['Service 1', 'Service 2', 'Service 3'], axis=1, inplace=True)
        
        with pd.ExcelWriter(r'C:\Users\Lenovo\Downloads\SidekickEDGE\Workbook1.xlsx') as writer:  
            new_df1.to_excel(writer, sheet_name='Sheet1', index=False)
            new_df2.to_excel(writer, sheet_name='Sheet2', index=False)
            new_df3.to_excel(writer, sheet_name='Sheet3', index=False)
            master.to_excel(writer, sheet_name='Master', index=False)
            
        new_dfx.to_excel(r'C:\Users\Lenovo\Downloads\SidekickEDGE\Workbook2.xlsx', sheet_name='Trends1')
        new_dfy.to_excel(r'C:\Users\Lenovo\Downloads\SidekickEDGE\Workbook3.xlsx', sheet_name='Trends2')       
        
    