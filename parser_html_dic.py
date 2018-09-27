from bs4 import BeautifulSoup
import pandas as pd

def get_table_title(soup):
    return soup.find('th').text

def get_columns_names(soup):
    nomes_colunas = [nome_coluna.text for nome_coluna in \
                     soup.find('tr', {'class': 'caption'}).findAll('th')]
    return nomes_colunas

def get_table_data(soup):
    rows = []
    for row in soup.find_all('tr', {'class': 'detail'}):
        linha_tabela = []
        for coluna in row.findAll('td'):
            linha_tabela.append(coluna.text)
        rows.append(linha_tabela)
        
    df = pd.DataFrame(data=rows, columns=get_columns_names(soup))
    #df.to_csv(dir_path + get_table_title(soup) + '.csv', sep=';', index=False)
    return df


def main():
    html1 = open('/home/cloves/Área de trabalho/TRTSP/dicionarios/sap1.html', 'r', encoding='latin-1') 
    html2 = open('/home/cloves/Área de trabalho/TRTSP/dicionarios/sap2.html', 'r', encoding='latin-1') 

    soup_sap1 = BeautifulSoup(html1.read(), 'html.parser')
    soup_sap2 = BeautifulSoup(html2.read(), 'html.parser')

    tabelas_sap1 = [tabela for tabela in soup_sap1.findAll('table')]
    tabelas_sap2 = [tabela for tabela in soup_sap2.findAll('table')]

    writer = pd.ExcelWriter('test.xlsx',engine='xlsxwriter')  
    workbook=writer.book
    cell_format = workbook.add_format({'bold': True, 'italic': False})
    worksheet=workbook.add_worksheet('SAP1')
    writer.sheets['SAP1'] = worksheet
    startrow = 1
    for tabela in tabelas_sap1:
        df = get_table_data(tabela)
        worksheet.write_string(startrow - 1, 0, get_table_title(tabela), cell_format)
        df.to_excel(writer,sheet_name='SAP1',startrow=startrow , startcol=0, index=False)
        startrow += df.shape[0] + 5

      
    worksheet=workbook.add_worksheet('SAP2')
    writer.sheets['SAP2'] = worksheet
    startrow = 1
    for tabela in tabelas_sap2:
        df = get_table_data(tabela)
        worksheet.write_string(startrow - 1, 0, get_table_title(tabela), cell_format)
        df.to_excel(writer,sheet_name='SAP2',startrow=startrow , startcol=0, index=False)
        startrow += df.shape[0] + 5

    writer.save()

if __name__=='__main__':
    main()