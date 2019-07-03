-- verificar espaco e uso de um database
USE BACEN;  
GO  
EXEC sp_spaceused @updateusage = N'TRUE';  
GO  

-- verificar espaco e uso de uma tabela
USE AdventureWorks2012;  
GO  
EXEC sp_spaceused N'Purchasing.Vendor';  
GO  