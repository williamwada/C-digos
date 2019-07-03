# -*- coding: utf-8 -*-
"""
Created on Fri Nov 25 09:50:54 2016

@author: Bruno.campos
"""
import pandas as pd
import numpy as np
import ctypes

def Mbox(title, text, style):
    ctypes.windll.user32.MessageBoxW(0, text, title, style)

def get_sidra(tabela,periodo,variavel,territorio,classificacoes):
    # http://api.sidra.ibge.gov.br/

    dadosSidra=pd.read_json('http://api.sidra.ibge.gov.br/values/t/'+tabela+'/p/'+periodo+'/v/'+variavel+territorio+classificacoes+'/f/c/d/m/h/n')

    return dadosSidra
    
########################################################################################################
    
parametros=pd.read_excel('//wilbor/C/Banco_de_dados_LCA/consultas_api/Sidra/consulta_sidra.xlsx',sheetname='consulta')
parametros=parametros.applymap(str)


sidraData=dict()
for i in range(0,len(parametros)):
    sidraData[parametros.iloc[i,1]]=get_sidra(parametros.iloc[i,2],parametros.iloc[i,3],parametros.iloc[i,4],parametros.iloc[i,5],parametros.iloc[i,6])
    sidraData[parametros.iloc[i,1]].drop('MC',axis=1,inplace=True)
    
    sidraData[parametros.iloc[i,1]]['V']=pd.to_numeric(sidraData[parametros.iloc[i,1]]['V'],errors='coerce')
    
    if parametros.frequencia[i]=='M':
        sidraData[parametros.iloc[i,1]]['D1C']=sidraData[parametros.iloc[i,1]]['D1C'].apply(lambda x: pd.to_datetime(x,format='%Y%m'))
        
    if parametros.frequencia[i]=='T':
        sidraData[parametros.iloc[i,1]]['D1C']=sidraData[parametros.iloc[i,1]]['D1C'].apply(lambda x: pd.to_datetime(int(x/100)*100+(x-int(x/100)*100)*3,format='%Y%m'))
        
    sidraData[parametros.iloc[i,1]]['column_code']=sidraData[parametros.iloc[i,1]].iloc[:,1].apply(str)
    sidraData[parametros.iloc[i,1]].drop(sidraData[parametros.iloc[i,1]].columns[1],axis=1,inplace=True)
    
    for j in range(2,len(sidraData[parametros.iloc[i,1]].columns)-1):
        sidraData[parametros.iloc[i,1]]['column_code']=sidraData[parametros.iloc[i,1]]['column_code']+'_'+sidraData[parametros.iloc[i,1]].iloc[:,1].apply(str)
        sidraData[parametros.iloc[i,1]].drop(sidraData[parametros.iloc[i,1]].columns[1],axis=1,inplace=True)
        
    sidraData[parametros.iloc[i,1]]=sidraData[parametros.iloc[i,1]].pivot(index='D1C',columns='column_code',values='V')

for destination in parametros.destination_file.unique():
    xlWriter = pd.ExcelWriter(destination.replace('\\','/'))
    for ws in parametros.loc[parametros.destination_file==destination,'destination_sheet']:
        sidraData[ws].to_excel(xlWriter,sheet_name=ws)
    xlWriter.save()
    
    Mbox('API Sidra', 'Consulta finalizada', 0)


