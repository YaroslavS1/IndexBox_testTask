import sqlite3
import pandas as pd
import numpy as np
import docx 

conn = sqlite3.connect("test.db") # connect to test.db 
cursor = conn.cursor() # get a cursor to the sqlite database 

# TODO все обернуть в класс и предоставить интерфейсы с помоью которых можно масштабировать логику

data = pd.read_sql('SELECT * FROM testidprod WHERE partner is NULL AND state is NULL AND bs = 0  AND factor in (1,2)', conn)
data = data.groupby('id').mean().reset_index()
data = data.drop(['mir', 'raw', 'bs', 'id', 'country'], axis=1, inplace=False)
data.loc[data['res'] == "NULL", 'res'] = np.nan
data = data.groupby(['factor', 'year']).sum().reset_index()


def add_year(data):
    year_new = [x for x in range(2006, 2021, 1)] # TODO добавить в конструктор 

    for i in year_new:
        if i not in data[lambda x: x['factor'] == 1]["year"].unique():
            # print(i)
            df = pd.DataFrame({'factor': [1], 'year': [int(i)], 'res': [np.nan]})
            # print(df)
            data = data.append(df, sort=True)
        if i not in data[lambda x: x['factor'] == 2]["year"].unique():
            # print(i)
            df = pd.DataFrame({'factor': [2], 'year': [int(i)], 'res': [np.nan]})
            # print(df)
            data = data.append(df, sort=True)

    data = data.sort_values(by=['factor', 'year'])
    col_one_list = data['res'].tolist()[:15]
    col_two_list = data['res'].tolist()[-15:]
    
    #! Добавляем в Dataframe factor 6
    df6 = pd.DataFrame()
    for i in range(len(year_new)):
        df = pd.DataFrame({'factor': [6], 'year': [year_new[i]], 'res': [col_one_list[i]/col_two_list[i]]})
        df6 = df6.append(df, sort=True)
    
    df6 = df6.reset_index(drop=True)
    df6 = df6.groupby(['factor', 'year']).mean().reset_index()
    #! Сохраняем Dataframe в .docx
    save_docx(df6)

    data = pd.concat([data, df6], ignore_index=True)
    data = data.reset_index(drop=True)
    data = data.groupby(['factor', 'year']).mean()
    data = data.rename(columns={'res': 'world'})
    data = pd.pivot_table(data, values = 'world', columns = ['factor', 'year'], dropna=False)
    print(data, '-')

    # TODO проверить сохранения в exel
    data.to_excel("report.xlsx")
    return data


def cagrf(df):
    """Function to calculate CAGR"""
    
    i_y = []
    for i in range(len(df['res'])):
        if str(df['res'][i]) != 'nan':
            i_y.append(i)
    return ((df['res'][i_y[-1]]/df['res'][i_y[0]])**(1/(df['year'][i_y[-1]]-df['year'][i_y[0]]))-1), df['year'][i_y[0]], df['year'][i_y[-1]]



# TODO сделать сохранение в WORD
def save_docx(df):
    doc = docx.Document()

    table = doc.add_table(rows = 16, cols = 3)
    cell = table.cell(1, 0)
    cell.merge(table.cell(15, 0))

    #добавляем заголовок таблицы
    table.cell(0,0).paragraphs[0].add_run('Factor').bold=True
    table.cell(0,1).paragraphs[0].add_run('Year').bold=True
    table.cell(0,2).paragraphs[0].add_run('World Value').bold=True

    cell = table.cell(1, 0)
    cell.text = str(df['factor'][0])

    for i in range(0,15):
        cell = table.cell(i+1, 1)
        cell.text = str(df['year'][i])
        cell = table.cell(i+1, 2)
        cell.text = str(df['res'][i])
    
    doc.add_paragraph()

    cagr, s_year, e_year = cagrf(df)
    if cagr > 0:
        doc.add_paragraph('Factor 6 grew  by avg {:.2%} every year from {} to {}.'.format(cagr, s_year, e_year))
    elif cagr < 0:
        doc.add_paragraph('Factor 6 decreased by avg {:.2%} every year from {} to {}.'.format(cagr, s_year, e_year))


    doc.save('table.docx')



# print(add_year(data))
add_year(data)

