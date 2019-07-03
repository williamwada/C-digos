SELECT a.*
FROM OPENROWSET('SQLNCLI', 'Server=Seattle1;Trusted_Connection=yes;',
     'SELECT GroupName, Name, DepartmentID
      FROM AdventureWorks2012.HumanResources.Department
      ORDER BY GroupName, Name') AS a;
GO


-- Utilizando OPENROWSET
--leitura de excel via SQL SERVER
--
SELECT * 
FROM OPENROWSET(
'Microsoft.ACE.OLEDB.12.0',
'Excel 12.0;Database=C:\Temp\Pasta1.xlsx',
[Planilha1$])




select *
from t_cotacao_paises
