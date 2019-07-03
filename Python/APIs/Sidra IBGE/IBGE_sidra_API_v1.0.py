# -*- coding: utf-8 -*-
"""
Created on Mon Apr 15 09:10:05 2019

@author: william.wada
"""
import requests #requisições para webservice
import pyodbc   #banco de dados
from requests.exceptions import HTTPError

#METODO Insercao dos parametros na URL para a Requisicao
def getURL(url,tabelaSidra,classificacaoLavouraSidra):
	url_ =url
	#inserindo a serie e o N no url de requisicao
	url_final =url_.format(tabelaSidra,classificacaoLavouraSidra)
	
	return url_final

#METODO Estabele Conexao    
def getConexao(server,database,username,password):
    #String de conexao
    conexao = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
        
    return conexao

#METODO para estruturar a query
def setQuery(query,num,i):
    #for i in range(0,num):
        #convert 103 corresponde ao formato dd/mm/yyyy            
        query_completa =query.format(address_data[i]['MC'],address_data[i]['D1N'],address_data[i]['D2N'],address_data[i]['D3N'],address_data[i]['D4N'],address_data[i]['MN'],address_data[i]['V'])
        print(query_completa)
        cursor.execute(query_completa)

#MAIN##########################################################################



#VARIAVEIS
tabelaSidra =1612
classificacaoLavouraSidra =81
url = "http://api.sidra.ibge.gov.br/values/h/n/t/{}/c{}/allxt/n3/all/p/last/v/214";

#dados para logar no banco de dados
server = 'WILLIAM-T\LCA_P1' #  named instance   
database = 'teste_P1' 
username = 'sistemaAPI' 
password = 'lca12345' 

# Criar a URL e fazer a Requisicao dos dados para web service
url_completa = getURL(url,tabelaSidra,classificacaoLavouraSidra)
#print(url_completa)
try:
    request = requests.get(url_completa) 
    address_data = request.json() 
    request.raise_for_status()
except HTTPError as http_err:
    print(f'Erro HTTP ocorreu: {http_err}')  # Python 3.6
except Exception as err:
    print(f'Outro erro ocorreu: {err}')  # Python 3.6
else:
    print('Requisicao feita com Sucesso.')    

  


# Acessar o banco de dados SQL Server#########################################    
#Fazer a conexao com o banco de dados
try:
    conexao = getConexao(server,database,username,password)
    cursor = conexao.cursor()
except Exception as e:
    print(e)
    print()
    print("Erro na tentativa de conexão com banco de dados.")

#Execucao da query no Banco de dados

#print(request.text)
#num=address_data.count()
num =len(address_data)
print('Print NUM')
print(num)

try:
   for i in range(0,num):    
       query ="INSERT INTO teste_API.dbo.t_producao (NN,D1N,D2N,D3N,D4N,MN,V) VALUES('{}','{}','{}','{}','{}','{}','{}')"          
       setQuery(query,num,i)
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
    print('Conexão Fechada.')
    