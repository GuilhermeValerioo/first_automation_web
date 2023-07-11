# importar bibliotecas
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pandas as pd
import time

# criando navegador
servico = Service(ChromeDriverManager().install())
nav = webdriver.Chrome(service=servico)

# importar/visualizar base de dados
tabela_produtos = pd.read_excel("buscas.xlsx")

def verificar_tem_termos_banidos(lista_termos_banidos, nome):
    tem_termos_banidos = False
    for palavra in lista_termos_banidos:
        if palavra in nome.lower():
            tem_termos_banidos = True
    return tem_termos_banidos

def verificar_todos_termos_produto(lista_termos_nome_produto, nome):
    tem_todos_termos_produto = True
    for palavra in lista_termos_nome_produto:
        if palavra not in nome.lower():
            tem_todos_termos_produto = False
    return tem_todos_termos_produto

def busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    
    # tratar os valores
    termos_banidos = termos_banidos.lower()
    produto = produto.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_nome_produto = produto.split(" ")
    lista_ofertas = []
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)
        
    # pesquisar no google
    nav.get("https://shopping.google.com.br/")
    nav.find_element(By.XPATH, '//*[@id="REsRA"]').send_keys(produto, '\ue007')

    # obter informações dos produtos da web
    lista_resultados = nav.find_elements(By.CLASS_NAME, 'i0X6df')

    for resultado in lista_resultados:
        nome = resultado.find_element(By.CLASS_NAME, 'tAxDx').text

        # analisar se contém termos banidos
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)

        # analisar se contém TODOS os termos do nome do produto
        tem_todos_termos_produto = verificar_todos_termos_produto(lista_termos_nome_produto, nome)

        # selecionar apenas os produtos válidos
        if not tem_termos_banidos and tem_todos_termos_produto:
            # tratamento do preço
            preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
            preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            preco = float(preco) 

            # condição do preço
            if preco_minimo <= preco <= preco_maximo:
                elemento_referencia = resultado.find_element(By.CLASS_NAME, 'bONr3b')
                elemento_pai = elemento_referencia.find_element(By.XPATH, '..')
                link = elemento_pai.get_attribute('href')
                lista_ofertas.append((nome, preco, link))
    return lista_ofertas

def busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    # tratar os valores
    termos_banidos = termos_banidos.lower()
    produto = produto.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_nome_produto = produto.split(" ")
    lista_ofertas = []
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)
    
    # buscar no buscape
    nav.get("https://www.buscape.com.br/")
    nav.find_element(By.XPATH, '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(produto,  '\ue007')
    
    #  obter resultados
    time.sleep(5)
    lista_resultados = nav.find_elements(By.CLASS_NAME, 'SearchCard_ProductCard_Inner__7JhKb')
    
    for resultado in lista_resultados:
        nome = resultado.find_element(By.CLASS_NAME, 'SearchCard_ProductCard_Name__ZaO5o').text
        nome = nome.lower()
        link = resultado.get_attribute("href")
        preco = resultado.find_element(By.CLASS_NAME, 'Text_MobileHeadingS__Zxam2').text
    
        # analisar se contém termos banidos
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)

        # analisar se contém TODOS os termos do nome do produto
        tem_todos_termos_produto = verificar_todos_termos_produto(lista_termos_nome_produto, nome)
        
        # analisar o preço
        if not tem_termos_banidos and tem_todos_termos_produto:
            preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            preco = float(preco) 
            if preco_minimo <= preco <= preco_maximo:
                lista_ofertas.append((nome, preco, link))
    
    # retornar a lista de ofertas do buscape
    return lista_ofertas

tabela_ofertas = pd.DataFrame()

for linha in tabela_produtos.index:
    
    # informações do produto na tabela
    produto = tabela_produtos.loc[linha, "Nome"]
    termos_banidos = tabela_produtos.loc[linha, "Termos banidos"]
    preco_minimo = tabela_produtos.loc[linha, "Preço mínimo"]
    preco_maximo = tabela_produtos.loc[linha, "Preço máximo"]

    # salvar as ofertas boas em um dataframe
    lista_ofertas_google_shopping = busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_google_shopping:
        tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns = ['produto', 'preço', 'link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_google_shopping])
    else:
        tabela_google_shopping = None
    lista_ofertas_buscape = busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns = ['produto', 'preço', 'link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_buscape])
    else:
        tabela_buscape = None

# exportar para o excel
tabela_ofertas.to_excel("Ofertas.xlsx", index = False)

# fechar navegador
nav.quit()