# -*- coding: utf-8 -*-
"""
Created on Mon Apr 15 09:10:05 2019

@author: william.wada
"""
import requests #requisições para webservice
from requests.exceptions import HTTPError
import pyodbc   #banco de dados
import logging
from datetime import date

### Criacao de log
data_atual =date.today()
logging.basicConfig(level=logging.DEBUG, filename='logs/IBGE_IPCA_'+str(data_atual)+'.log', filemode='a', format='%(process)d - %(asctime)s - %(name)s - %(levelname)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S')
##


### METODO Insercao dos parametros na URL para a Requisicao
def getURL(url,tabelaSidra,classificacaoLavouraSidra):
	url_ =url
	#inserindo a serie e o N no url de requisicao
	url_final =url_.format(tabelaSidra,classificacaoLavouraSidra)
	
	return url_final
##	

###METODO Estabele Conexao    
def getConexao(server,database,username,password):
#String de conexao
	conexao = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
	return conexao
##

###METODO para estruturar a query
def setQuery(query,num,i):
		#convert 103 corresponde ao formato dd/mm/yyyy
		query_completa =query.format(
			#address_data[i]['D1C'].replace("'","''").translate ({ord(c): "" for c in "¨!@#$%^&*()[]{};:/<>?\|`~=_+"}),
			address_data[i]['D1C'].replace("'", "''"),
			address_data[i]['D1N'].replace("'","''"),
			address_data[i]['D2N'].replace("'","''"),
			address_data[i]['D3N'].replace("'","''"),
			address_data[i]['D4N'].replace("'","''"),
			address_data[i]['MN'].replace("'","''"),
			address_data[i]['V'].replace("'","''")
		)
		#print(query_completa)
		cursor.execute(query_completa)

##

###MAIN##########################################################################

#VARIAVEIS
#tabelaSidra =1612
#classificacaoLavouraSidra =81
url_completa = "http://api.sidra.ibge.gov.br/values/t/1419/n7/all/v/63/p/201701/c315/all/f/u/h/n";

#dados para logar no banco de dados

dado_url_comum ="C:/Users/william.wada/Desktop/OneDrive - LCA Consultores/Códigos/Python/APIs/Sidra IBGE/names/"

#url      = open(dado_url_comum+str('ibge_ipca_url.txt')).read().strip()
server   = open(dado_url_comum+str('ibge_ipca_sqlserver_server.txt')).read().strip()
database = open(dado_url_comum+str('ibge_ipca_sqlserver_database.txt')).read().strip()
username = open(dado_url_comum+str('ibge_ipca_sqlserver_username.txt')).read().strip()
password = open(dado_url_comum+str('ibge_ipca_sqlserver_password.txt')).read().strip()

url_completa = "http://api.sidra.ibge.gov.br/values/t/1419/n7/all/v/63/p/201701/c315/all/f/u/h/n";

# Criar a URL e fazer a Requisicao dos dados para web service
#url_completa = getURL(url,tabelaSidra,classificacaoLavouraSidra)
#print(url_completa)

#Verificacao se conectou com sucesso na API
try:
	#url_completa = getURL(url,ser,num)
	logging.warning('Fazendo requisição ao Web Service.')
	request = requests.get(url_completa)
	address_data = request.json()
	logging.warning('Requisição de dados realizado com sucesso.')
	request.raise_for_status()
except HTTPError as http_err:
	print(f'Erro HTTP ocorreu: {http_err}')  # Python 3.6
	logging.exception("Erro HTTP ocorreu: {http_err}")
except Exception as err:
	print(f'Outro erro ocorreu: {err}')  # Python 3.6
	logging.exception("Outro erro ocorreu: {err}")
else:
	print('Requisicao feita com Sucesso.')

# Acessar o banco de dados SQL Server#########################################    
#Fazer a conexao com o banco de dados
try:
	logging.warning('Conectando na base de dados.')
	conexao = getConexao(server,database,username,password)
	logging.warning('Conectado com a base de dados.')
	cursor = conexao.cursor()
except Exception as e:
	print(e)
	print()
	print("Erro na tentativa de conexão com banco de dados.")
	logging.exception("Erro na tentativa de conexão com banco de dados.")


###Execucao da query no Banco de dados
	
num =len(address_data)

try:
	logging.warning('Executando query na base de dados.')
	for i in range(0,num):
		query ="INSERT INTO teste_API.dbo.t_stage_IPCA (D1C,D1N,D2N,D3N,D4N,MN,V) VALUES('{}','{}','{}','{}','{}','{}','{}')"
		setQuery(query,num,i)
	conexao.commit()
	print('Inserção com Sucesso na base de dados.')
	logging.warning('Query executada com sucesso na base de dados.')
except Exception as e: 
	conexao.rollback()
	print('Erro na execução da query, rollback efetuado.')
	print()
	print(e)
	logging.exception("Erro na execução da query, rollback efetuado.")

finally:        
	conexao.close()
	print()
	print('Conexão Fechada.')
	logging.warning("Conexão com banco Finalizada.")
##
	
	
	
