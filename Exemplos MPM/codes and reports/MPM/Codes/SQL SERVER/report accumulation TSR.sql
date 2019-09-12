USE [EPOS]
GO
/****** Object:  StoredProcedure [dbo].[KDDI_TSR_Report_Acc]    Script Date: 11/02/2015 12:53:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Ry Jiang>
-- Create date: <31 Mar 2015>
-- Description:	<KDDI TSR Accumulation report>
-- =============================================
ALTER PROCEDURE [dbo].[KDDI_TSR_Report_Acc] 
AS
BEGIN

	declare @tgtDir varchar(255)
	declare @tmpFile varchar(255) 
	declare @tgtFile varchar(255) 
    declare @cmd varchar(2000)

	--------------------------------------------------
	declare @CurrentDate datetime
	declare @CurrentYear varchar(4)
	declare @CurrentMonth varchar(2)
	--The year that report is created
	declare @ReportYear varchar(4)
	--The month that report is created
	declare @ReportMonth varchar(2)
	--Report name
	declare @ReportName varchar(50)
	--The command to create the report
	declare @SqlSelTSRReport varchar(200)
	--The result of searching the report records in the table TSR_REPORT_ACC 
	declare @RowCount int
	--The number of report being created(for example, if the reportcount=2, means the reports for this month and last month will be created)
	declare @ReportCount int
	--This parameter is to show whether the report is created
	declare @ReportRecord int 
	--For the current date
	declare @currentdateYYYYMMDD varchar(10)
	-------------[Start] Added by Jiang on 2014/03/03------------------------
	declare @LoadDateCount int
	declare @ReportDate date
	declare @MaxContactDate varchar(8)
	declare @DirDate varchar(8)
	-------------[End]--------------------------
	set @currentdate = GETDATE()
	--set @currentdate = dateadd(day,-1,GETDATE())

	--Change the format for the current date to YYYY/MM/DD
	set @currentdateYYYYMMDD = convert(varchar,@currentdate,112)
	--The current year
	set @CurrentYear = YEAR(@currentdate) 
	--The current month
	set @CurrentMonth = MONTH(@currentdate)
	--The template file name
	set @tmpFile = 'D:\EPOS\Reports\Template\TSR_Accumulation_Report.xls'
	--To set the report name
	set @ReportName = 'EPOSIB_TSR_ACCUMULATION_REPORT' 
	set @ReportCount = 0
	set @ReportRecord = 0
	set @RowCount = 0
    select @DirDate = convert(varchar,MAX(CONTACT_DATE_TIME),112) from CallDetailRecord
	set @tgtDir = 'D:\EPOS\Reports\TSR_Accumulation_Report\ADCF_TSR_Accumulation_Report_' + @DirDate + '\'
	set @cmd = 'mkdir ' + @tgtDir
	EXEC master..xp_cmdshell @cmd, no_output
	----------------------------------------------------------------------------------
	--If the report record of this month in the TSR_REPORT_MONTHLY table does not exists,
	--   1)Create the report for this month
	--   2)Check whether the report record of last month in the TSR_REPORT_MONTHLY table exists
    ---------If yes, do nothing(end the loop)
    ---------else,   create the report for the last month 	       
	--Else if the report record of this month in the TSR_REPORT_ACC table exists 
	--   1)Create the report for this month, then end the loop.
	while (@ReportCount < 3)
	begin
		-- Initialize TMP TABLE
		delete from TMP_TSR_REPORT

		set @ReportYear = year(DATEADD(m, - @ReportCount, @currentdate))
	    set @ReportMonth = MONTH(DATEADD(m, - @ReportCount, @currentdate))

	        ---------Read TSR_REPORT_ACC TABLE------------------------------------
			declare CPCursor cursor for
			select COUNT(*) 
			from TSR_REPORT_ACC 
			where REPORT_NAME = @ReportName 
				  and REPORT_YEAR = + @ReportYear  
				  and  REPORT_MONTH = @ReportMonth
			
			open CPCursor
			
			fetch next from CPCursor
			into @RowCount

			if (@RowCount = 0)
				---------The report for this month does not exist----------------
				set @ReportRecord = 0
			else
			    -------	The report for this month exist,	
				set @ReportRecord = 1
			---------------------------------
			if @ReportRecord = 1 and @ReportCount > 0
			begin
				-----[Start]Modified by Jiang on 2014/03/03
				set @ReportDate = cast(@ReportYear + '-' + @ReportMonth + '-01' as DATE)
				select @LoadDateCount = COUNT(*)
				from
				(
						select distinct convert(date,DATEADD(mm,datediff(mm,0,CONTACT_DATE_TIME),0),111) as clldate
						from CallDetailRecord 
						where load_date = CONVERT(date,@currentdate,111)
						----------[Start]Modified on 2014-06-05--------
						/*union
						select distinct convert(date,DATEADD(mm,datediff(mm,0,CONTACT_DATE_TIME),0),111) as clldate
						from OCS.dbo.CallDetailRecord
						where load_date = CONVERT(date,@currentdate,111)*/
						/*select distinct convert(date,DATEADD(mm,datediff(mm,0,CONTACT_DATE_TIME),0),111) as clldate
						from TCC.dbo.CallDetailRecord
						where load_date = CONVERT(date,@currentdate,111)*/
						----------[End]2014-06-05----------------------
				) t
				where t.clldate = @ReportDate
				
				if(@LoadDateCount = 0 )
				Begin
				-----[End]-------------------------------
					close CPCursor
					deallocate CPCursor
					break
				-----[Start] Added by Jiang on 2014/03/03---------------------------
				End
				-----[End]-----------------------------	
			end
			close CPCursor
			deallocate CPCursor
			------------------------------Create Report-------------------------------
			begin
			----------------create report--------------------------------------------
 				select @MaxContactDate = convert(varchar,MAX(CALL_DATE),112)
				from DAILY_TSR_REPORT 
				where year(CALL_DATE) = @ReportYear and MONTH(CALL_DATE)= @ReportMonth

				--Copy the template file
				set @tgtFile = @tgtDir + 'EPOSIB_TSR_Accumulation_Report_' + @MaxContactDate +'.xls'		
				set @cmd = 'COPY /Y '+ @tmpFile + ' /B ' + @tgtFile

				EXEC master..xp_cmdshell @cmd, no_output
			    --Select data from DB Epos and TCC, Put the data into the temp table
			    insert into TMP_TSR_REPORT (TSR_ID_NAME,CALL_DATE,
							OB_LHOURS,OB_THOURS,
							OB_THOURSP,OB_RHOURS,OB_RHOURSP,OB_WHOURS,OB_WHOURSP,
							OB_DIALS,OB_DIALSP,OB_CONNECTS,OB_CONNECTSP,OI_DMC,OI_CPH,OI_FDMC,OI_FC,
							OB_SALE1,IB_SALE1,OB_TTL_SALE1,IB_TTL_SALE1,OI_SPC,OI_SPH,OB_MP1,IB_MP1,
							OI_MPPH,OB_SALE2,IB_SALE2,OI_DMCP,OI_CAUTIONS,OB_THOUR_CONNECT,
							OB_THOUR_SALE,CAMPAIGN_CODE,IB_THOURS,IB_CONNECTS)
				
			    select   m.TSR_REPORT_KID + ',' + m.TSR_NAME,CALL_DATE, 
							OB_LHOURS,OB_THOURS,
							OB_THOURSP,OB_RHOURS,OB_RHOURSP,OB_WHOURS,OB_WHOURSP,
							OB_DIALS,OB_DIALSP,OB_CONNECTS,OB_CONNECTSP,OI_DMC,OI_CPH,OI_FDMC,OI_FC,
							--------[Start]Modified by R.J. on 2014/08/05
							--OB_SALE1,IB_SALE1,OB_TTL_SALE1,IB_TTL_SALE1,OI_SPC,OI_SPH,OB_MP1,IB_MP1,
							OB_TTL_SALE1,IB_TTL_SALE1,OB_TTL_SALE1,IB_TTL_SALE1,OI_SPC,OI_SPH,OB_MP1,IB_MP1,
							--------[End]---------------------------							
							OI_MPPH,OB_SALE2,IB_SALE2,OI_DMCP,OI_CAUTIONS,OB_THOUR_CONNECT,
							OB_THOUR_SALE,CAMPAIGN_CODE,coalesce(IB_THOURS,0) as IB_THOURS,coalesce(IB_CONNECTS,0) as IB_CONNECTS 
				from DAILY_TSR_REPORT, TSR_MASTER_KDDI m 
				where year(CALL_DATE) = @ReportYear and MONTH(CALL_DATE)= @ReportMonth and DAILY_TSR_REPORT.TSR_ID_NAME = m.TSR_ID
				---------------[Start]Added on 2014/06/05
				/*union
	            select   TSR_ID_NAME + ',' + m.TSR_NAME,CALL_DATE, 
							OB_LHOURS,OB_THOURS,
							OB_THOURSP,OB_RHOURS,OB_RHOURSP,OB_WHOURS,OB_WHOURSP,
							OB_DIALS,OB_DIALSP,OB_CONNECTS,OB_CONNECTSP,OI_DMC,OI_CPH,OI_FDMC,OI_FC,
							OB_SALE1,IB_SALE1,OB_TTL_SALE1,IB_TTL_SALE1,OI_SPC,OI_SPH,OB_MP1,IB_MP1,
							OI_MPPH,OB_SALE2,IB_SALE2,OI_DMCP,OI_CAUTIONS,OB_THOUR_CONNECT,
							OB_THOUR_SALE,CAMPAIGN_CODE,coalesce(IB_THOURS,0) as IB_THOURS,coalesce(IB_CONNECTS,0) as IB_CONNECTS 
				from OCS.dbo.DAILY_TSR_REPORT, OCS.dbo.TSR_MASTER m 
				where year(CALL_DATE) = @ReportYear and MONTH(CALL_DATE)= @ReportMonth and DAILY_TSR_REPORT.TSR_ID_NAME = m.TSR_ID*/			
				/*
				union 
				select   TSR_ID_NAME + ',' + m.TSR_NAME ,CALL_DATE,
							 OB_LHOURS,OB_THOURS,
							 OB_THOURSP,OB_RHOURS,OB_RHOURSP,OB_WHOURS,OB_WHOURSP,
							 OB_DIALS,OB_DIALSP,OB_CONNECTS,OB_CONNECTSP,OI_DMC,OI_CPH,OI_FDMC,OI_FC,
							 OB_SALE1,IB_SALE1,OB_TTL_SALE1,IB_TTL_SALE1,OI_SPC,OI_SPH,OB_MP1,IB_MP1,
							 OI_MPPH, OB_SALE2,IB_SALE2,OI_DMCP,OI_CAUTIONS,OB_THOUR_CONNECT,
							 OB_THOUR_SALE,CAMPAIGN_CODE,
							 --null as IB_THOURS,null as IB_CONNECTS 
							 0 as IB_THOURS,0 as IB_CONNECTS
				from TCC.dbo.TSR_ACTIVITY_REPORT r, TCC.dbo.TSR_MASTER_KDDI m 
				where year(CALL_DATE) = @ReportYear and 
							 MONTH(CALL_DATE)= @ReportMonth and r.TSR_ID_NAME = m.TSR_ID */
				---------------[End]---------------------
				
				--Put the data into excel file			 
				set @cmd = 'insert into openrowset(''Microsoft.ACE.OLEDB.12.0''' +
							',''Excel 8.0;Database=' + @tgtFile +
							''',''select * from [Data$]'')' +  
						   ' select ROW_NUMBER() OVER (ORDER BY  call_date,TSR_ID_NAME,CAMPAIGN_CODE )AS ROWID,  TSR_ID_NAME ' + 
							  ',CONVERT(varchar,CALL_DATE,111) as call_date' + 
							  ',OB_LHOURS,OB_THOURS,OB_THOURSP,OB_RHOURS,OB_RHOURSP,OB_WHOURS' +
							  ',OB_WHOURSP,OB_DIALS,OB_DIALSP,OB_CONNECTS,OB_CONNECTSP' +	
							  ',OI_DMC,OI_CPH,OI_FDMC,OI_FC,OB_SALE1,IB_SALE1,OB_TTL_SALE1' + 	
							  ',IB_TTL_SALE1,OI_SPC,OI_SPH,OB_MP1,IB_MP1,OI_MPPH' + 	
							  ',OB_SALE2,IB_SALE2,OI_DMCP,OI_CAUTIONS,OB_THOUR_CONNECT' + 	
							  ',OB_THOUR_SALE,CAMPAIGN_CODE,IB_THOURS,IB_CONNECTS' +	
								' from TMP_TSR_REPORT' +
								' where year(CALL_DATE) = ' + @ReportYear +  
								' and MONTH(CALL_DATE)= ' + @ReportMonth +
								' order by call_date'
	
				--execute sp_executesql @cmd 
				exec sp_sqlexec @cmd
			-----------------------------------------------------------------------
	
				
			----------Insert or update TSR_REPORT_MONTHLY tbl-----------------------
			--If the report record does not exist in the table TSR_REPORT_MONTHLY
			--   insert 
			--Else
			--   update the user name and the update date time 
			if @ReportRecord = 0
				Insert into TSR_REPORT_ACC (REPORT_NAME, REPORT_YEAR,REPORT_MONTH, REPORT_USER_NAME, CREATE_DATE,UPDATE_DATE_TIME) 
				Values (@ReportName,@ReportYear, @ReportMonth, SUSER_SNAME(), convert(date,@CurrentDate,111), @CurrentDate )
			else
				update	TSR_REPORT_ACC 
				set UPDATE_DATE_TIME = @CurrentDate, REPORT_USER_NAME = SUSER_SNAME()
				where REPORT_NAME = @ReportName and REPORT_YEAR = @ReportYear and REPORT_MONTH = @ReportMonth
			-------------------------------------------------------------------------
			
			end
			--------------------------------------------------------------------------
			set @ReportCount = @ReportCount + 1;
	--End the loop
	end
	 	 
END

