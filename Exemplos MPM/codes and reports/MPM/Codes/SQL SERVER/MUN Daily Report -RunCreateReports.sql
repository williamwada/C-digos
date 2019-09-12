USE [DDS_MUNNK_PRD]
GO
/*
select campaign_code, count(*) from dds_cdr
where insert_date_time >= cast(getdate() as date)
group by campaign_code
*/

DECLARE @return_value int

-- TODO: Set parameter values here.

EXECUTE @return_value = [dbo].[MUN_Summary_Report] 

if @return_value=0
select 'MUN_Summary_Report' = 'Completed'
else
select 'MUN_Summary_Report' = 'Failed'

EXECUTE @return_value = [dbo].[MUN_KPI_Report] 

if @return_value=0
select 'MUN_KPI_Report' = 'Completed'
else
select 'MUN_KPI_Report' = 'Failed'

EXECUTE @return_value = [dbo].[MUN_Daily_Sales_Report] 

if @return_value=0
select 'MUN_Daily_Sales_Report' = 'Completed'
else
select 'MUN_Daily_Sales_Report' = 'Failed'

EXECUTE @return_value = [dbo].[MUN_SALES_CUSTOMERS] 

if @return_value=0
select 'MUN_SALES_CUSTOMERS' = 'Completed'
else
select 'MUN_SALES_CUSTOMERS' = 'Failed'

GO

