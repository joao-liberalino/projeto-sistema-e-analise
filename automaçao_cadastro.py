# criar um projeto ded automaçao que :
# abre o windows 
# digitar o nome do arquivo de cadastro (projetodisbra - Atalho)
# clicar nele 
# cadastrar os clientes das planilhas e pular para proxima area 


#bibliotecas 
import pyautogui as pya 
import time
import pandas as pd 


#entra no sitema 
pya.PAUSE = 1 #da um tempo nos comando 
pya.press("win") #abrie windows
pya.write("projetodisbra - Atalh") #digitar busca
pya.click(x=763, y=395)
pya.press("enter")
time.sleep(2)


#ler base de dados 

tabela = pd.read_json("C:/Users/joaov/Downloads/cadastros_ficticios_sp.json")
pd.set_option('display.max_columns', None)  # Mostra todas as colunas
pd.set_option('display.width', None)        # Não limita a largura
pd.set_option('display.max_colwidth', None)
print(tabela)

#cadastrar OS PRODUTOS 
pya.press("tab")

for linha in tabela.index:
    CNPJ = str(tabela.loc[linha,"CNPJ"])
    pya.write(CNPJ)

    pya.press("tab")
    NOME_CLIEN = str(tabela.loc[linha,"NOME_CLIEN"])
    pya.write(NOME_CLIEN)

    pya.press("tab")
    FONE = str(tabela.loc[linha,"FONE/FAX"])
    pya.write(FONE)

    pya.press("tab")
    CEP = str(tabela.loc[linha,"CEP"])
    pya.write(CEP)

    pya.press("tab")
    ENDERECO = str(tabela.loc[linha,"ENDEREÇO"])
    pya.write(ENDERECO)

    pya.press("tab")
    BAIRRO = str(tabela.loc[linha,"BAIRRO"])
    pya.write(BAIRRO)

    pya.press("tab")
    MUNICIPIO = str(tabela.loc[linha, "MUNICIPIO"])
    pya.write(MUNICIPIO)

    pya.press("tab")
    UF = str(tabela.loc[linha,"UF"])
    pya.write(UF)

    pya.press("tab")
    DATA_EMISSAO = str(tabela.loc[linha, "DATA_EMISSAO"])
    pya.write(DATA_EMISSAO)

    pya.press("tab")
    TANQUE = str(tabela.loc[linha, "TANQUE"])
    pya.write(TANQUE)

    pya.press("tab")
    QUANTIDADE = str(tabela.loc[linha, "QUANTIDADE"])
    pya.write(QUANTIDADE)

    pya.press("tab")
    BOMBA = str(tabela.loc[linha, "BOMBA"])
    pya.write(BOMBA)

    pya.press("tab")
    MODELO = str(tabela.loc[linha, "MODELO"])
    pya.write(MODELO)

    pya.press("tab")
    BACIA = str(tabela.loc[linha, "BACIA"])
    pya.write(BACIA)

    pya.press("tab")
    pya.click(x=230, y=442)
    pya.click(x=230, y=442)
    pya.click(x=230, y=442)
    pya.click(x=230, y=442)
    pya.click(x=230, y=442)    
    pya.click(x=230, y=442)
    pya.click(x=230, y=442)
    pya.click(x=230, y=442)
    pya.click(x=230, y=442)
    pya.click(x=230, y=442)
    pya.click(x=230, y=442)
    pya.click(x=230, y=442)
    pya.click(x=230, y=442)
    pya.click(x=230, y=442)
    

    
    






    





