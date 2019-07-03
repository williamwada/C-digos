-- Acessando o banco de dados onde desejo a pesquisa
 
USE [BACEN]
GO
 
-- Iniciando a pesquisa nas tabelas de sistemas
 
SELECT A.NAME, A.TYPE, B.TEXT
  FROM SYSOBJECTS  A (nolock)
  JOIN SYSCOMMENTS B (nolock) 
    ON A.ID = B.ID
WHERE B.TEXT LIKE '%TB_AUX_MERC_MEDIANA_EXPORT%'  --- Informação a ser procurada no corpo da procedure, funcao ou view
 -- AND A.TYPE = 'P'                     --- Tipo de objeto a ser localizado no caso procedure
 ORDER BY A.NAME
 
GO