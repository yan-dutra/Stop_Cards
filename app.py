#Importa as bibliotecas necessárias
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt 
import seaborn as sns
from datetime import datetime, timedelta
import base64
import io
from PIL import Image

#pywin32==302

#Titúlo
image = Image.open('logo1.jpg')
st.image(image, width=350)
st.header('**Sistema de Tratamento e Análise de Dados do Formistica**')
st.write("Esse sistema viza facilitar o processo de automatização da transcrição dos Cartões de Desvio para o sistema da Petrobras\n e também faz uma análise de dados nos cartões.")

################################### ETL ###################################
    
#Ler o arquivo do usuário
st.markdown('---')
arq = st.file_uploader('Faça o upload do arquivo .xlsx:', type='.xls')

#Processa o arquivo
if arq:
    st.markdown('---')
    st.subheader('**Aqui está o arquivo tratado 🤖**')
    df = pd.read_excel(arq)
        
    #Filtra as colunas necessárias
    df = df.filter(items = ["Gerência/ Empresa do local da ocorrência:", "AddedUserDescription", "Data da ocorrência",
    "Hora da ocorrência","Unidade operacional/Sonda", "Classificação do desvio/ Categoria do evento:", "Desvio Sistêmico", "Desvio Crítico",
    "Regras de Ouro Segurança - Petrobras", "O que aconteceu?", "Ação imediata", "Local da ocorrência"])

    #Adiciona 5 minutos na coluna "Hora da ocorrência e
    # salva na nova coluna "Fim da ocorrência"
    for h in df['Hora da ocorrência']:
        fim_ocorrencia = df['Hora da ocorrência'] + timedelta(minutes=5)
        df['Fim da ocorrência'] = fim_ocorrencia

    #Converte o formato timedelta64 para string
    df['Fim da ocorrência'] = df['Fim da ocorrência'].astype(str).map(lambda x: x[7:])

    #Sepera a data
    df['time'] = df['Data da ocorrência'].str.split().apply(pd.Series)
    df['datetime'] = pd.to_datetime(df['time'])
    df['Dia'] = df['datetime'].dt.day
    df['Mês'] = df['datetime'].dt.month
    df['Ano'] = df['datetime'].dt.year
    df = df.drop(['time', 'datetime'], 1)

    #Separa a hora
    df['Hora_Início'] =  df['Hora da ocorrência'].str[:2]           
    df['Minuto_Início'] =  df['Hora da ocorrência'].str[3:5]
    df['Hora_Fim'] =  df['Fim da ocorrência'].str[:2]
    df['Minuto_Fim'] =  df['Fim da ocorrência'].str[3:5]


    #Organiza as colunas
    df = df[["Gerência/ Empresa do local da ocorrência:", "AddedUserDescription", "Data da ocorrência", "Hora da ocorrência", "Fim da ocorrência",
    "Unidade operacional/Sonda", "Classificação do desvio/ Categoria do evento:", "Desvio Sistêmico", "Desvio Crítico", 
    "Regras de Ouro Segurança - Petrobras", "O que aconteceu?", "Ação imediata", "Local da ocorrência", "Dia", "Mês", "Ano", "Hora_Início", 
    "Minuto_Início", "Hora_Fim", "Minuto_Fim"]]

    #Mostra o resultado 
    st.markdown('\n')
    st.dataframe(df)

    #Faz o download do arquivo
    st.markdown('---')
    st.markdown('**Faça o download:**')
    towrite = io.BytesIO()
    downloaded_file = df.to_excel(towrite, encoding='utf-8', index=False, header=True)
    towrite.seek(0)  # redefine o ponteiro
    b64 = base64.b64encode(towrite.read()).decode()  
    linko= f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="base_tratada.xlsx">Download da base tratada</a>'
    st.markdown(linko, unsafe_allow_html=True)

        

    ################################### INSIGHTS ###################################

    st.markdown('---')
    st.subheader("**E temos alguns insights 🧐**")

    #Cria uma coluna 'Cartão' que recebe o valor 1 para cada elemento
    #na coluna 'AddedUserDescription'
    for n in df['AddedUserDescription']:
        df['Cartão'] = 1
        
    #Total de Cartões no Mês
    st.markdown('\n')
    total_c = df['Cartão'].value_counts(dropna=False)
    st.write("Total de Cartões no Mês:", total_c)

    #Cartões por Técnico
    st.markdown('---')
    cartao_tecnico = df[['AddedUserDescription', 'Cartão']]
    cartao_tecnico = cartao_tecnico.groupby('AddedUserDescription', dropna=False).count().reset_index()
    cartao_tecnico = cartao_tecnico.sort_values(by='Cartão', ascending=False)

    sns.catplot(x='AddedUserDescription', y='Cartão', palette='mako',
                kind='bar',
                data=cartao_tecnico,
                height=5,
                aspect=2)

    plt.axhline(14, 0, 1, color='crimson', **{'ls':'--'})
    plt.xticks(fontsize = 10, rotation=90)
    plt.style.use('ggplot')
    plt.title('Cartões por Técnico')
    plt.show()
    st.pyplot(plt)
    plt.clf()

    st.markdown('\n')
    st.write("***Obs:***", "Será que os ténicos que não atingiram a meta foi devido terem embarcado perto do mês virar? ")


    #Cartões por Cliente
    st.markdown('---')
    cartao_cliente = df[['Gerência/ Empresa do local da ocorrência:', 'Cartão']]
    cartao_cliente = cartao_cliente.groupby('Gerência/ Empresa do local da ocorrência:', dropna=False).count().reset_index()
    cartao_cliente = cartao_cliente.sort_values(by='Cartão', ascending=False)

    sns.catplot(x='Gerência/ Empresa do local da ocorrência:', y='Cartão', palette='mako',
                kind='bar',
                data=cartao_cliente,
                height=5,
                aspect=2)

    plt.xticks(fontsize = 10, rotation=90)
    plt.style.use('ggplot')
    plt.title('Cartões por Cliente')
    plt.show()
    st.pyplot(plt)
    plt.clf()


    #Cartões por Sonda
    st.markdown('---')
    cartao_sonda = df[['Unidade operacional/Sonda', 'Cartão']]
    cartao_sonda = cartao_sonda.groupby('Unidade operacional/Sonda', dropna=False).count().reset_index()
    cartao_sonda = cartao_sonda.sort_values(by='Cartão', ascending=False)

    sns.catplot(x='Unidade operacional/Sonda', y='Cartão', palette='mako',
                kind='bar',
                data=cartao_sonda,
                height=5,
                aspect=2)

    plt.xticks(fontsize = 10, rotation=90)
    plt.style.use('ggplot')
    plt.title('Cartões por Sonda')
    plt.show()
    st.pyplot(plt)
    plt.clf()


    #Cartões por Local de Ocorrência
    st.markdown('---')
    cartao_local = df[['Local da ocorrência', 'Cartão']]
    cartao_local = cartao_local.groupby('Local da ocorrência', dropna=False).count().reset_index()
    cartao_local = cartao_local.sort_values(by='Cartão', ascending=False)

    sns.catplot(x='Local da ocorrência', y='Cartão', palette='mako',
                kind='bar',
                data=cartao_local,
                height=5,
                aspect=2)

    plt.xticks(fontsize = 10, rotation=90)
    plt.style.use('ggplot')
    plt.title('Cartões por Local de Ocorrência')
    plt.show()
    st.pyplot(plt)
    plt.clf()


    #Cartões por Categoria do Evento
    st.markdown('---')
    cartao_categoria = df[['Classificação do desvio/ Categoria do evento:', 'Cartão']]
    cartao_categoria = cartao_categoria.groupby('Classificação do desvio/ Categoria do evento:', dropna=False).count().reset_index()
    cartao_categoria = cartao_categoria.sort_values(by='Cartão', ascending=False)

    sns.catplot(x='Classificação do desvio/ Categoria do evento:', y='Cartão', palette='mako',
                kind='bar',
                data=cartao_categoria,
                height=5,
                aspect=2)

    plt.xticks(fontsize = 10, rotation=90)
    plt.style.use('ggplot')
    plt.title('Cartões por Categoria do Evento')
    plt.show()
    st.pyplot(plt)
    plt.clf()


    #Cartões por Regra de Ouro
    st.markdown('---')
    cartao_regra = df[['Regras de Ouro Segurança - Petrobras', 'Cartão']]
    cartao_regra = cartao_regra.groupby('Regras de Ouro Segurança - Petrobras', dropna=False).count().reset_index()
    cartao_regra = cartao_regra.sort_values(by='Cartão', ascending=False)

    sns.catplot(x='Regras de Ouro Segurança - Petrobras', y='Cartão', palette='mako',
                kind='bar',
                data=cartao_regra,
                height=5,
                aspect=2)

    plt.xticks(fontsize = 10, rotation=90)
    plt.style.use('ggplot')
    plt.title('Cartões por Regras de ouro da Petrobras')
    plt.show()
    st.pyplot(plt)
    plt.clf()
        

    #Cartões por Desvio Sistêmico
    st.markdown('---')
    cartao_ds = df[['Desvio Sistêmico', 'Cartão']]
    cartao_ds = cartao_ds.groupby('Desvio Sistêmico', dropna=False).count().reset_index()
    cartao_ds = cartao_ds.sort_values(by='Cartão',ascending=False)

    sns.catplot(x='Desvio Sistêmico', y='Cartão', palette='mako',
                kind='bar',
                data=cartao_ds,
                height=5,
                aspect=2)

    plt.xticks(fontsize = 12, rotation=0)
    plt.style.use('ggplot')
    plt.title('Cartões por Desvio Sistêmico')
    plt.show()
    st.pyplot(plt)
    plt.clf()


    #Cartões por Desvio Crítico
    st.markdown('---')
    cartao_dc = df[['Desvio Crítico', 'Cartão']]
    cartao_dc = cartao_dc.groupby('Desvio Crítico', dropna=False).count().reset_index()
    cartao_dc = cartao_dc.sort_values(by='Cartão',ascending=False)

    sns.catplot(x='Desvio Crítico', y='Cartão', palette='mako',
            kind='bar',
            data=cartao_dc,
            height=5,
            aspect=2)

    plt.xticks(fontsize = 12, rotation=0)
    plt.style.use('ggplot')
    plt.title('Cartões por Desvio Crítico')
    plt.show()
    st.pyplot(plt)
    plt.clf()
    st.balloons()