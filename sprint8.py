#Sobre: Manipulando planilhas do Google, com o Google Sheets API e utilizando Python.

from logging import error
import gspread
from numpy.core.defchararray import index
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import time
from sqlalchemy import create_engine
import smtplib
import smtplib
import smtplib
import email.message
import time

start_time = time.time()

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

def send_email(frase, tempo):
    arq = open(r"C:\Users\lucas\Desktop\Mike\python\NappAcademy\api\\senha.txt") 
    senha = arq.readlines() #senha armazena em txt local 

    corpo_email = f'''<p>{frase}<p>
    <p>Tempo de execução {round(tempo,5)} segundos<p>''' #E-mail para ser enviado

    msg = email.message.Message()
    msg['Subject'] = 'API - GOOGLE'
    msg['From'] = 'mike.william@nappsolutions.com' #Remetente
    msg['To'] = 'mike.william@nappsolutions.com'    #destinatário
    password = f'{str(senha[0])}' #Senha do Email
    msg.add_header('Content-Type', 'text/html') 
    msg.set_payload(corpo_email )
    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')


def importa_DB(df):

    #df = pd.read_csv(r'C:\Users\lucas\Desktop\Mike\Tarefas\002- API\soma_produto_in_estados.csv', sep = ';', encoding='utf-8') - Exemplo importação de csv pra DF

    df.columns = [c.lower() for c in df.columns] #postgres não aceita maiúsculas ou espaços

    from sqlalchemy import create_engine #biblioteca de importação
    engine = create_engine('postgresql://postgres:154878@localhost:5432/testedb') #CONNECTION STRING

    df.to_sql("napp_academy", engine) #Definindo nome da tabela no banco

    print('Importado para o banco local!')

#Função Main
def main():
    frase = ''
    try: 
        df = planilha_to_df()
        
        
        #print(df.shape) 122 x 13
        df = transforma(df)


        importa_DB(df)
        print(df)

        frase = 'SUCESSO ao executar script!'
    except:
        frase = 'ERRO ao executar script!'
    end_time = time.time()

    tempo = end_time - start_time

    send_email(frase, tempo)

    print(frase)
main()


#Fontes de estudo:
#https://ichi.pro/pt/como-ativar-o-acesso-do-python-ao-planilhas-google-250853149332571 - chave
#https://pt.linkedin.com/pulse/manipulando-planilhas-do-google-usando-python-renan-pessoa - configurando script
#https://www.hashtagtreinamentos.com/enviar-email-gmail-python - send-email

#Criar um projeto Google Cloud Plataform: https://console.developers.google.com/project
#Link das Planilhas: https://docs.google.com/spreadsheets/u/0/