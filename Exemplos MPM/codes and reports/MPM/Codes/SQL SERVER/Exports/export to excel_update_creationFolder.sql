USE [DDS_EPOS_PRD]
GO
/****** Object:  StoredProcedure [dbo].[DAM_EposDaily_Report]    Script Date: 05/30/2017 11:05:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<W W>
-- Create date: <11 April 2017>
-- Description:	<Web Epos Sales Daily report>
-- Who        When             What
-- 
-- =============================================
ALTER procedure [dbo].[DAM_EposDaily_Report]
@File_qtd int -- file qtd
AS
BEGIN
DECLARE @SalesStartdate varchar(20)
DECLARE @cnt int
SET @cnt=@File_qtd

WHILE @cnt > 0
BEGIN
   SET @SalesStartdate=GETDATE()- @cnt;
  
--Clear 
DELETE TM_WebEposReport where sales_date = @SalesStartdate

INSERT TM_WebEposReport
	select * 
	from
	(	
	SELECT dt.mydate as sales_date,
	ISNULL(MGI_Data.Sales_Count_MGI,0) as Sales_Count_MGI,
	ISNULL(CAN_Data.Sales_Count_Can,0) as Sales_Count_CAN,
	ISNULL(MGI_Data.MP_MGI,0) as MP_MGI,
	ISNULL(CAN_Data.MP_CAN,0) as MP_CAN	
	FROM DatesTable(convert(varchar,@SalesStartdate,111),convert(varchar,@SalesStartdate,111)) as dt
	left join
	(	
	--MGI
	select F13.DATE_SUBSCRIPTION,COUNT(*) as Sales_Count_MGI,NULL as Sales_Count_Can,SUM(MP) as MP_MGI,NULL as MP_CAN
	from dbo.DAM_EPOS_F13 F13
	where CATEGORIES ='M'
	group by F13.DATE_SUBSCRIPTION 
	)as MGI_Data on dt.mydate = MGI_Data.DATE_SUBSCRIPTION
	left join
	(
	--CAN
	select F13.DATE_SUBSCRIPTION,NULL as Sales_Count_MGI,COUNT(*) as Sales_Count_Can,NULL as MP_MGI,SUM(MP) as MP_CAN
	from dbo.DAM_EPOS_F13 F13
	where CATEGORIES ='C'
	group by F13.DATE_SUBSCRIPTION 
	) as CAN_Data
	on dt.mydate=CAN_Data.DATE_SUBSCRIPTION
) as WebReport

SET @cnt = @cnt - 1;  
   
END;

SET @SalesStartdate=GETDATE()-1;
	
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
	
	set @TemplateName = 'WEB_EPOS_Sales'

	--make folder for the report
	set @tgtDir = 'D:\shared\DAM_EPOS\reports\'+@TemplateName+'_' + @rpdate +'\'
	set @cmd = 'mkdir ' + @tgtDir
	--print @cmd
	EXEC master..xp_cmdshell @cmd, no_output
	
	set @tmpFile = 'D:\EPOS_PRD\REPORTS\Template\'+@TemplateName+'.xls'
		set @tgtFile = @tgtDir + @TemplateName+'_'+@rpdate+'.xls'
		set @cmd = 'COPY /Y '+ @tmpFile + ' /B ' + @tgtFile
		print @cmd
		-- copy template file to daily folder
		EXEC master..xp_cmdshell @cmd, no_output
		
	set @SQL = 'update openrowset(''Microsoft.ACE.OLEDB.12.0''' +
			' ,''Excel 8.0;Database=' + @tgtFile + '''' + 

			' ,''select * from [data$]'')' + 	
			' set salesdate = dt.sales_date,mgiCount = dt.Sales_Count_MGI,'+
			' canCount = dt.Sales_Count_CAN,mgiMP =dt.MP_MGI,canMP=dt.MP_CAN'+
			' from TM_WebEposReport as dt where dt.sales_date =SALESDATE'
			exec sp_sqlexec @sql
		
END



