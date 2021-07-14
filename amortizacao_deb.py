# -*- coding: utf-8 -*-
"""
Web-scraping das datas de pagamento de juros e amortizações das debêntures do site da Anbima Datausando Selenium e Beautiful Soup

@author: Marcos
"""

import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
from bs4 import BeautifulSoup

#____________________________________________________________________________________________________
# Funções
def get_amortizacao(i):
    
    driver.get(df.iloc[i][1]) #Acessa a url
    driver.maximize_window() #Maximiza a tela
    driver.implicitly_wait(10) #Espera carregar a página

    #Pega o número de páginas
    test = driver.find_element_by_class_name("anbima-ui-pagination__right") #Acessa div das páginas
    test_span = test.find_elements_by_tag_name('span') #Acessa a tag de informações sobre as páginas
    last_page = test_span[1].text #Acessa o valor total de páginas
    last_page = int(last_page)

    lista_temp = [] #Lista que recebe as informações das debêntures
    
    time.sleep(1)
    
    for i in range(last_page):
    
        #Se ta na primeira página
        if (i == 0):
            table = driver.find_element_by_class_name("table__agenda") #Localiza a tabela dentro da página
            table_tbody = table.find_element_by_tag_name('tbody') #Localiza o tbody dentro da classe table agenda
            table_tr = table_tbody.find_elements_by_tag_name('tr') #Localiza as tags tr dentro da tabela
            
            for elemento in table_tr:
                card = {}
                temp = elemento.text.split("\n")
                card['Data do Evento'] = temp[0]
                card['Data da Liquidação'] = temp[1]
                card['Evento'] = temp[2]
                card['Percentual/Taxa (%)'] = temp[3]
                card['Status'] = temp[4]
                card['Valor Pago (R$)'] = temp[5]
                    
                lista_temp.append(card)
            
            time.sleep(0.5)
    
        #Passa de página
        if (i+1) > 1:
            input_element = driver.find_element_by_class_name("agenda-pagination") #Acessa o form de escrita
            input_element_ = input_element.find_elements_by_tag_name('input') #Acessa o bloco onde inputa o valor da página requerida
            
            input_element_[0].send_keys(Keys.LEFT) #Leva o cursor pra esquerda
            input_element_[0].send_keys(Keys.DELETE) #Deleta o valor atual
            input_element_[0].send_keys(i+1) #Reescreve para acessar a página desejada
            input_element_[0].send_keys(Keys.ENTER) #Aperta ENTER
            
            time.sleep(2)
            
            page = driver.page_source #Pega o html da página
            source = BeautifulSoup(page, 'html.parser') #Transforma numa variável BeautifulSoup
            table = source.find(class_ = 'table__agenda')
            table_tbody = table.tbody
            table_tr = table_tbody.findAll('tr') 
            
            for elemento in table_tr:
                card = {}
                temp = elemento.findAll('td')
                card['Data do Evento'] = temp[0].text
                card['Data da Liquidação'] = temp[1].text
                card['Evento'] = temp[2].text
                card['Percentual/Taxa (%)'] = temp[3].text
                card['Status'] = temp[4].text
                card['Valor Pago (R$)'] = temp[5].text
                
                #print(card)
                lista_temp.append(card)
                
            time.sleep(0.5)
            
    return lista_temp


#____________________________________________________________________________________________________
# Leitura das urls
df = pd.read_excel('D:/2020/Data Science/Alura/web-scraping-deb/deb_url.xlsx', index_col=0)

urls = df.iloc[:,1]

#Substitui a string 'características' por 'agenda' nas url's
for url in range(len(urls)):
    n = urls.iloc[url].split(sep='/')
    n[-1] = 'agenda'
    df.iloc[url][1] = '/'.join(n)  
    
#____________________________________________________________________________________________________
#Autentica acesso ao Chrome
chromedriver = r'C:/Users/Marcos/Downloads/chromedriver'
driver = webdriver.Chrome(executable_path=chromedriver)

lista = [] #Lista que recebe as debêntures

#Iteração main
for i in range(0, len(df)):
    try:
        temp = get_amortizacao(i)
        lista.append(temp)
        print(df.iloc[i][0] + "- " + str(i) + "/" + str(len(df) - 1))
        
    except:
        print("erro: "+ df.iloc[i][0])
        continue
    
#Salva a lista em um dataframe
df1 = pd.DataFrame(lista)

#Altera o index para o nome da debênture
df1 = df1.reset_index()
for i in range(0,len(df)):
    df1['index'][i] = df.iloc[i][0]

#Exporta para o Excel
df1.to_excel('amortizacao_deb.xlsx')