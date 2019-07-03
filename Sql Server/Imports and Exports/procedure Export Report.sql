USE [teste_API]
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<W W>
-- Create date: <16 April 2019>
-- Description:	<testando criacao de novo excel via template>
--
-- Who        When             What
-- 
-- =============================================
ALTER procedure [dbo].[BC_test_Report]
@File_qtd int -- file qtd
AS
BEGIN
DECLARE @Startdate varchar(20)
DECLARE @cnt int
SET @cnt=@File_qtd

WHILE @cnt > 0
BEGIN
   SET @Startdate=GETDATE();

--
INSERT TMP_t_cotacao_paises
select * 
from teste_API.dbo.t_cotacao_paises order by serie   



SET @cnt = @cnt -1;
END

-- pump to Excel
	declare @cmd varchar(1000)
	declare @tgtDir varchar(255)
	declare @tmpFile varchar(255) 
	declare @tgtFile varchar(255) 
	declare @TemplateName varchar(20)
	declare @CPstartdt date
	declare @CPenddt date
	declare @rpdate varchar(20)
	declare @SQL varchar(2000)
	
	select @rpdate = CONVERT(varchar,GETDATE()-1,112) 
	
	set @TemplateName = 'TMP_testing'

	--make folder for the report
	set @tgtDir = 'C:\Temp\relatorio\'+@TemplateName+'_' + @rpdate +'\'
	set @cmd = 'mkdir ' + @tgtDir
	--print @cmd
	EXEC master..xp_cmdshell @cmd, no_output
	
	set @tmpFile = 'C:\Temp\'+@TemplateName+'.xlsx'
		set @tgtFile = @tgtDir + @TemplateName+'_'+@rpdate+'.xlsx'
		set @cmd = 'COPY /Y '+ @tmpFile + ' /B ' + @tgtFile
		print @cmd
		-- copy template file to daily folder
		EXEC master..xp_cmdshell @cmd, no_output
		
	set @SQL = 'update openrowset(''Microsoft.ACE.OLEDB.12.0''' +
			' ,''Excel 12.0;Database=' + @tgtFile + '''' + 

			' ,''select * from [Planilha1$]'')' + 	
			' set serie = dt.serie,valor = dt.valor,'+
			' data = dt.data,pais =dt.pais'+
			' from TMP_t_cotacao_paises as dt where dt.data =data'

		print @SQL
			begin try
				exec sp_sqlexec @sql
			end try
			begin catch
			-- Execute error retrieval routine.  
			
				print 'Error na sp_sqlexec'
			end catch
END



exec [dbo].[BC_test_Report] 1;






INSERT INTO OPENROWSET('Microsoft.ACE.OLEDB.12.0',  
    'Excel 12.0;Database=C:\Temp\testing.xlsx;',  
    'SELECT * FROM [Planilha1$]') select* from [teste_API].[dbo].[t_cotacao_paises]

	select * from [teste_API].[dbo].[t_cotacao_paises]


	select * from TMP_t_cotacao_paises

	update openrowset('Microsoft.ACE.OLEDB.12.0' ,'Excel 12.0;Database=C:\Temp\relatorio\TMP_testing_20190415\TMP_testing_20190415.xlsx' ,'select * from [Planilha1$]') set serie1 = dt.serie,valor1 = dt.valor, data1 = dt.data,pais1 =dt.pais from TMP_t_cotacao_paises as dt





