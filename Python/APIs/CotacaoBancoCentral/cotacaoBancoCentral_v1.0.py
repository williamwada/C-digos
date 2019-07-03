# -*- coding: utf-8 -*-
"""
Created on Mon Apr 15 09:10:05 2019

@author: william.wada
"""
import requests #requisições para webservice
import pyodbc   #banco de dados


#inputs da moeda e periodo #10813
#serie_input =input('Digite a serie que deseja consultar:')
#n_input =int(input('Digite os N ultimos valores desejado:  '))

def getCotacao(serie,n):
	url ='http://api.bcb.gov.br/dados/serie/bcdata.sgs.{}/dados/ultimos/{}?formato=json'
	#inserindo a serie e o N no url de requisicao
	url_final =url.format(serie,n)
	
	return url_final


url_completa = getCotacao('10813',100)


print(url_completa)
print()
# Requisicao dos dados
request = requests.get(url_completa)

#Requisicao no formato json
address_data = request.json()


# Acessar o banco de dados SQL Server######################
    
#dados para logar no banco de dados
server = 'WILLIAM-T\LCA_P1' #  named instance
database = 'teste_P1' 
username = 'sistemaAPI' 
password = 'lca12345' 


#String de conexao
try:
    conexao = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
    cursor = conexao.cursor()
except Exception as e:
    print(e)
    print()
    print("Erro na tentativa de conexão com banco de dados")



try:
    for i in range(0,n_input):    
       #convert 103 corresponde ao formato dd/mm/yyyy
        query ="INSERT INTO teste_API.dbo.t_cotacao_paises (serie,valor, data) VALUES("+str(i+1)+",{},CONVERT(DATE,'{}',103))"
        query_completa =query.format(address_data[i]['valor'],address_data[i]['data'])
       # print(query_completa)
        
        cursor.execute(query_completa)
    conexao.commit()
    print('Inserção com Sucesso na base de dados.')
        
except Exception as e: 
    conexao.rollback()
    print('Erro na execução da query, rollback efetuado.')
    print()
    print(e)
    
        
finally:
    conexao.close()
    print()
    print('Conexão Fechada')
    
