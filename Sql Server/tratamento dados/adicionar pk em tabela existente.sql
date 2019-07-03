USE AdventureWorks2012;  
GO  
ALTER TABLE Production.TransactionHistoryArchive   
ADD CONSTRAINT PK_TransactionHistoryArchive_TransactionID PRIMARY KEY CLUSTERED (TransactionID);  
GO  