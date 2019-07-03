
--1 Export data to existing EXCEL file from SQL Server table  
INSERT INTO OPENROWSET('Microsoft.ACE.OLEDB.12.0',  
    'Excel 12.0;Database=C:\Temp\testing.xlsx;',  
    'SELECT * FROM [Planilha1$]') select* from [dbo].[t_cotacao_paises]
  
  
--2 Export data from Excel to new SQL Server table  
SELECT * 
INTO [dbo].[t_cotacao_paises_from_excel] FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',  
    'Excel 12.0;Database=C:\Temp\testing.xlsx;HDR=YES',  
    'SELECT * FROM [Planilha1$]')  
  
  
--3 Export data from Excel to existing SQL Server table  
INSERT INTO [dbo].[t_cotacao_paises_from_excel] Select * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',  
   'Excel 12.0;Database=C:\Temp\testing.xlsx;HDR=YES',  
   'SELECT * FROM [Planilha1$]')  
  
  
--4 If you dont want to create an EXCEL file in advance and want to export data to it, use  
  
EXEC sp_makewebtask  
    @outputfile = 'c:\Temp\testingWithoutTemplate.xls',   
    @query = 'Select * from [teste_API].[dbo].[t_cotacao_paises_from_excel]',   
    @colheaders = 1,   
    @FixedFont = 0,@lastupdated = 0,@resultstitle = 'Testing details'  
 --   (Now you can find the file with data in tabular format)  
  
  
--5 To export data to new EXCEL file with heading(column names), create the following procedure  
  
CREATE PROCEDURE proc_generate_excel_with_columns  
(  
    @db_name    varchar(100),  
    @table_name varchar(100),  
    @file_name  varchar(100)  
)  
as  
  
--Generate column names as a recordset  
declare @columns varchar(8000), @sql varchar(8000), @data_file varchar(100)  
select  
    @columns = coalesce(@columns + ',', '') + column_name + ' as ' + column_name  
from  
    information_schema.columns  
where  
    table_name = @table_name  
select @columns = '''''' + replace(replace(@columns, ' as ', ''''' as '), ',', ',''''')  
  
--Create a dummy file to have actual data  
select @data_file = substring(@file_name, 1, len(@file_name) - charindex('\',reverse(@file_name)))+'\data_file.xls'  
  
--Generate column names in the passed EXCEL file  
set @sql = 'exec master..xp_cmdshell ''bcp " select * from (select ' + @columns + ') as t" queryout "' + @file_name + '" -c'''  
exec(@sql)  
  
--Generate data in the dummy file  
set @sql = 'exec master..xp_cmdshell ''bcp "select * from ' + @db_name + '..' + @table_name + '" queryout "' + @data_file + '" -c'''  
exec(@sql)  
  
--Copy dummy file to passed EXCEL file  
set @sql = 'exec master..xp_cmdshell ''type ' + @data_file + ' >> "' + @file_name + '"'''  
exec(@sql)  
  
--Delete dummy file  
set @sql = 'exec master..xp_cmdshell ''del ' + @data_file + ''''  
exec(@sql)  
  
--After creating the procedure, execute it by supplying database name, table name and file path  
  
--EXEC proc_generate_excel_with_columns 'your dbname', 'your table name', 'your file path'  
EXEC proc_generate_excel_with_columns 'teste_API', 't_cotacao_paises', 'C:\Temp\excelDaProcedure.xlsx'  
