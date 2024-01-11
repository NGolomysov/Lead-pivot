import openpyxl
import pandas
import pandas as pd
import plotly.express as px
import numpy as np
import streamlit as st
import datetime
import time
import base64
from io import StringIO, BytesIO

st.set_page_config(page_title='Воронка клиентов',
                   page_icon=':page_facing_up:',
                   layout='wide')

uploaded_file = st.file_uploader("Загрузите файл")

def generate_excel_download_link(df):
    towrite = BytesIO()
    df.to_excel(towrite, encoding="ANSI", index=False, header=True)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_download.xlsx">Скачать таблицу</a>'
    return st.markdown(href, unsafe_allow_html=True)


@st.cache_data                                     #cash
def get_data():
    df = pd.read_excel(uploaded_file)
    df.rename(columns={"Id лида": "Id",
                       "Дата поступления": "Date",
                       "Источник поступления": "Source",
                       "ФИО лида": "LeadName",
                       "Статус лида": "Status",
                       "Код курса": "Code",
                       "ФИО менеджера": "MngName"
                       }, inplace=True)

    df['Source'] = df['Source'].fillna('Пустые')
    df['Date'] = pd.to_datetime(df['Date'])
    return df

try:

    df = get_data()

    #Sidebar

    code = st.sidebar.multiselect(
        'Выберите коды курсов для отчёта:',
        options=df['Code'].unique(),
        default=df['Code'].unique()
    )

    mngname = st.sidebar.multiselect(
        'Выберите менеджеров для отчёта:',
        options=df['MngName'].unique(),
        default=df['MngName'].unique()
    )

    df = df.query(
        'Code == @code & MngName == @mngname'
    )

    df['total_by_source'] = (df.groupby(['Source'])['Id'].transform(lambda x: len(x.unique())))

    leads = df.groupby(['Source', 'Status']).size().reset_index(name="by_source")

    leads_piv = pd.pivot_table(leads,
                               index = ['Source'],
                               columns = 'Status',
                               values = 'by_source'
                               ).fillna('0').astype('int')

    report = pd.DataFrame(leads_piv).reset_index().fillna(0)

    border_stamp = df.iloc[df["Date"].argmax()]['Date']
    border = pd.to_datetime(str(border_stamp.month) + '/' + '01/' + str(border_stamp.year))

    report = report.merge(
        df.groupby('Source')['Date'].apply(lambda x: (x >= border).sum()).reset_index(name='За месяц'),
        how='outer')
    try:
        report['Брак/дубль'] = report['Брак'].astype('int')+report['Дубль'].astype('int')
        report = report.drop(columns=['Брак', 'Дубль'])
    except KeyError as err:
        try:
            if err.args[0] == 'Дубль':
                report['Брак/дубль'] = report['Брак'].astype('int')
                report = report.drop(columns=['Брак'])
            else:
                report['Брак/дубль'] = report['Дубль'].astype('int')
                report = report.drop(columns=['Дубль'])
        except KeyError:
            report['Брак/дубль'] = 0



    report_clients = df.loc[df['Status'] == 'Стал клиентом'].reset_index(drop = True)

    report = report.merge(
        report_clients.groupby('Source')['Date'].apply(lambda x: (x >= border).sum()).reset_index(name='Стали клиентами в этом месяце'),
        how='outer')
    report['Стали клиентами в этом месяце'] = report['Стали клиентами в этом месяце'].fillna(0).astype('int')

    report = report.merge(
        report_clients.groupby('Source')['Date'].apply(lambda x: (x < border).sum()).reset_index(name='Стали клиентом с прошлого периода'),
        how='outer')
    report['Стали клиентом с прошлого периода'] = report['Стали клиентом с прошлого периода'].fillna(0).astype('int')

    report = report.merge(
        report_clients.groupby('Source')['total_by_source'].count().reset_index(name='Всего клиентов'),
        how='outer')
    report['Всего клиентов'] = report['Всего клиентов'].fillna(0).astype('int')

    report = report.merge(
        df.groupby('Source')['total_by_source'].count().reset_index(name='Всего'),
        how='outer')
    report['Всего'] = report['Всего'].fillna(0).astype('int')

    col_order = ['Источник поступления',
                 'Всего',
                 'За месяц',
                 'Новый',
                 'Горячий',
                 'Теплый',
                 'Холодный',
                 'Брак/дубль',
                 'Лист ожидания',
                 'Стали клиентами в этом месяце',
                 'Стали клиентом с прошлого периода',
                 'Всего клиентов']


    report = report.rename(columns={"Source": "Источник поступления"})
    real_cols = list(report.columns.values)
    cols = [x for x in col_order if x in real_cols]
    indexes = []
    for el in cols:
        indexes.append(real_cols.index(el))

    report = report.iloc[:,indexes]
    #report = pd.concat([report.sum().rename('Всего'), report]).reset_index(drop = True)
    report = report.append(report.sum().rename('Всего'))
    report.at['Всего', 'Источник поступления'] = 'Всего'

    st.dataframe(report)

    generate_excel_download_link(report)
except:
    st.error('Введите данные')


#HIDE STREAMLIT STYLE
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)
