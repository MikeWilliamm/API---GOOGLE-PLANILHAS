#Sobre: Manipulando planilhas do Google, com o Google Sheets API e utilizando Python.

import gspread
from numpy.core.defchararray import index
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import numpy as np
import os as o
from pandas.core.arrays import integer
import psycopg2 
import time
from sqlalchemy import create_engine
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib, ssl

#Acessa planilha online e transforma em um Data Frame
def planilha_to_df(): 
#Escopo utilizado
    scope = ['https://spreadsheets.google.com/feeds']

    #Dados de autenticação
    credentials = ServiceAccountCredentials.from_json_keyfile_name(r'C:\Users\lucas\Desktop\Mike\python\NappAcademy\api\nappacademy-854d5535b91f.json', scope)

    # #Se autentica
    gc = gspread.authorize(credentials)

    # #Abre a planilha
    wks = gc.open_by_key('1nv1w0jGpBvYQqYCqWD1OjWFOhThoPg1qsyIDEyzxx7c')

    #Seleciona a primeira página da planilha
    worksheet = wks.get_worksheet(0)

    #Transforma planilha em DataFrame
    df = pd.DataFrame(worksheet.get_all_records())
    
    #print(df)
    print('Planilha transformada em DataFrame')
    return df

#Transformação de dados
def transforma(df): 
    df['statusLimite'] = 0 #Cria uma nova coluna chamada statusLimite
    df = df.values #Transforma o DataFrame em uma lista

    media = 0
    for i in range(len(df)): #Calcula o valor médio de limites de crédito
        media += df[i,12]
    media = media / len(df)
    
    #Se o valor do limite for maior que o valor médio então a coluna 'statusimite' recebe 'Limite ALTO'
    #Se o valor do limite for menor que o valor médio então a coluna 'statusimite' recebe 'Limite BIXO'
    for i in range(len(df)): 
        if df[i,12] > media:
            df[i,13] = 'Limite ALTO'
        else:
            df[i,13] = 'Limite BAIXO'
        print(df[i])

    lista_colunas = ['customerNumber', 'customerName', 'contactLastName', 'contactFirstName', 'phone', 'addressLine1', 'addressLine2', 'city', 'state', 'postalCode', 'country', 'salesRepEmployeeNumber', 'creditLimit', 'statusLimit']    
    df = pd.DataFrame(df, columns= lista_colunas) #Transforma a lista em um DataFrame novamente
    print(df)
    
    return df


def importa_DB(df):

    #df = pd.read_csv(r'C:\Users\lucas\Desktop\Mike\Tarefas\002- API\soma_produto_in_estados.csv', sep = ';', encoding='utf-8') - Exemplo importação de csv pra DF

    df.columns = [c.lower() for c in df.columns] #postgres não aceita maiúsculas ou espaços

    from sqlalchemy import create_engine #biblioteca de importação
    engine = create_engine('postgresql://postgres:154878@localhost:5432/testedb') #CONNECTION STRING

    df.to_sql("napp_academy", engine) #Definindo nome da tabela no banco

    print('Importado para o banco local!')

#Função Main
def main():
    df = planilha_to_df()
    
    
    #print(df.shape) 122 x 13
    df = transforma(df)


    importa_DB(df)
    #print(df)
main()


#Fontes de estudo:
#https://ichi.pro/pt/como-ativar-o-acesso-do-python-ao-planilhas-google-250853149332571 - chave
#https://pt.linkedin.com/pulse/manipulando-planilhas-do-google-usando-python-renan-pessoa - configurando script

#Criar um projeto Google Cloud Plataform: https://console.developers.google.com/project
#Link das Planilhas: https://docs.google.com/spreadsheets/u/0/