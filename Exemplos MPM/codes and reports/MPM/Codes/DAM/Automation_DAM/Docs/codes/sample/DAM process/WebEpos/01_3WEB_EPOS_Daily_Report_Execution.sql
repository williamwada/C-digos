

/***
	Daily reports excution
*/
USE EPOS 
GO

DECLARE	@return_value int

EXEC	@return_value = [dbo].[WebEposDaily_Report]
if @return_value = 0
SELECT	'WebEpos_Daily_Report' = 'Completed!'
else
SELECT	'WebEpos_Daily_Report' = 'Failed!'