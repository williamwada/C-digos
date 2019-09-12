USE [EPOS]
GO

/****** Object:  StoredProcedure [dbo].[WebEposDaily_Report]    Script Date: 04/10/2017 16:34:55 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<W W>
-- Create date: <25 December 2015>
-- Description:	<Web Epos Sales Daily report>
-- Who        When             What
-- 
-- =============================================
CREATE procedure [dbo].[WebEposDaily_Report] as

BEGIN
DECLARE @SalesStartdate varchar(20)
SET @SalesStartdate='2015-11-17';
DROP TABLE TM_WebEposReport;	
	select * into TM_WebEposReport from
	(
	SELECT dt.mydate as sales_date,
	ISNULL(TMGI.Sales_Count,0) as Sales_Count_MGI,
	ISNULL(TCAN.Sales_Count,0) as Sales_Count_CAN ,
	ISNULL(TMGI.MP,0) as MP_MGI,
	ISNULL(TCAN.MP_CAN,0) as MP_CAN
	FROM 
	DatesTable(convert(varchar,@SalesStartdate,111),convert(varchar,GETDATE(),111)) as dt
	left join 
	(
	--MGI
	select 'M' as type,convert(date,convert(varchar,SALE_DATE)) as SALE_DATE,
		   SUM(Total_Premium) as MP,
		   SUM(BASe.Sales_count1+BASe.Sales_count2) as Sales_Count
		from dbo.WEB_EPOS_MGI_YES_FILE A,
		(
		  select Customer_id,
		   SUM(CASE when Main_Premium <>'' THEN 1 ELSE 0 END) as Sales_count1,
		   SUM(CASE when Spouse_Premium <>'' THEN 1 ELSE 0 END) as Sales_count2
		  from dbo.WEB_EPOS_MGI_YES_FILE
		  group by Customer_id
		)BASe
		where A.CUSTOMER_ID =BASe.CUSTOMER_ID
	group by SALE_DATE
	) as TMGI on  dt.mydate = TMGI.SALE_DATE
	left join
	(
	--CAN
	select convert(date,convert(varchar,SALE_DATE)) as SalesDate,
		SUM(Main_Premium) as MP_CAN,
		SUM(BASe.Sales_count1+BASe.Sales_count2) as Sales_Count
		from dbo.WEB_EPOS_CAN_YES_FILE A,
		(
		  select Customer_id,
		   SUM(CASE when Main_Premium <>'' THEN 1 ELSE 0 END) as Sales_count1,
		   SUM(CASE when Spouse_Flag <>'' THEN 1 ELSE 0 END) as Sales_count2
		  from dbo.WEB_EPOS_CAN_YES_FILE
		  group by Customer_id
		)BASe
		where A.CUSTOMER_ID =BASe.CUSTOMER_ID
		group by convert(date,convert(varchar,SALE_DATE))
	) as TCAN on dt.mydate = TCAN.SalesDate
	)as WebReport
	
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
	
	set @TemplateName = 'EPOS_WEB_Sales'

	--make folder for the report
	set @tgtDir = 'D:\EPOS\Reports\WebEpos\'+@TemplateName+'_' + @rpdate +'\'
	set @cmd = 'mkdir ' + @tgtDir
	EXEC master..xp_cmdshell @cmd, no_output
	
	set @tmpFile = 'D:\EPOS\Reports\Template\'+@TemplateName+'.xls'
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



GO


