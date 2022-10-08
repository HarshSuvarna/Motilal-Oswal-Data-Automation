import pandas as pd
import xlsxwriter
import os
import glob
import pandas as pd


path = os.getcwd()
files = glob.glob(os.path.join(path, '*.xlsx'))
csv_files = glob.glob(os.path.join(path, '*csv'))
files.append(csv_files[0])

for file in files:
    
    if file.split('.')[1] == 'xlsx':
        data = pd.read_excel(file)
        file_name = file.split('\\')[-1][:-5]
    else:
        data = pd.read_csv(file)
        file_name = file.split('\\')[-1][:-4]
    
    

    data.insert(21, 'NETVALUEa', abs(data['NETVALUE']))

    gb_broker = data.groupby('FASUBTYPECODE')[['NETVALUEa']].sum()
    gb_broker.reset_index(inplace=True)
    gb_broker.sort_values(by='NETVALUEa', inplace=True, ascending=False)
    gb_broker.rename(columns={'NETVALUEa':'Sum of NETVALUE'}, inplace=True)
    gb_broker.loc[len(gb_broker.index)] = ['Grand Total', gb_broker['Sum of NETVALUE'].sum()]

    gb_company = data.groupby('SCRIPNAME')[['NETVALUE']].sum()
    gb_company.reset_index(inplace=True)
    gb_company.sort_values(by='NETVALUE', inplace=True, ascending=False)
    gb_company.rename(columns={'NETVALUE':'Sum of NETVALUE'}, inplace=True)
    gb_company.loc[len(gb_company.index)] = ['Grand Total', gb_company['Sum of NETVALUE'].sum()]

    #file_name = file.split('\\')[-1][:-5]
    all_sheets = {file_name:data, 'Broker data':gb_broker, 'Company data':gb_company}

    f = file.split('\\')[0:-1]
    f = '//'.join(f)
    export_path = f+'\\OP'+file_name+'.xlsx'

    writer = pd.ExcelWriter(export_path, engine = 'xlsxwriter')

    data.drop(columns='NETVALUEa', inplace=True)

    for sheet_name in all_sheets.keys():
        all_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)

    writer.save()
    
