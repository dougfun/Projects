# Final script

import pandas as pd
import os
import datetime


data = datetime.datetime.now()

# Creating Dataframe with final structure (of consolidated)
columns = [
    'Segmento',
    'País',
    'Produto',
    'Qtde de Unidades Vendidas',
    'Preço Unitário',
    'Valor Total',
    'Desconto',
    'Valor Total c/ Desconto',
    'Custo Total',
    'Lucro',
    'Data',
    'Mês',
    'Ano'
]

consolidado = pd.DataFrame(columns = columns)

# Searches the name of the files to be consolidated
files = os.listdir("planilhas")

# Makes the Consolidated files (excel only .xlsx)
for excel in files:
    if excel.endswith('.xlsx'):
        data_file = excel.split('_')
        segmento = data_file[0]
        pais = data_file[1].replace('.xlsx', '')

        # This will create a dataframe with the content we analysed
        try:
            df = pd.read_excel(f'planilhas\\{excel}')
            df.insert(0, 'Segmento', segmento)
            df.insert(1, 'País', pais)
            consolidado = consolidado.append(df)
            
        except:
            with open('log_errors.txt', 'w') as file:
                file.write(f'Error trying to consolidate the file {excel}.')
    else:
        with open('log_errors.txt', 'w') as file:
            file.write(f'The file {excel} is not valid as excel.')


consolidado['Data'] = pd.to_datetime(consolidado['Data']) 
consolidado['Data'] = consolidado['Data'].dt.strftime('%d/%m/%Y')

# Exports the consolidated dataframe to excel 
consolidado.to_excel(f"Report-consolidado-{data.strftime('%d-%m-%Y')}.xlsx", 
                     index = False,
                     sheet_name = 'Report Consolidated')
