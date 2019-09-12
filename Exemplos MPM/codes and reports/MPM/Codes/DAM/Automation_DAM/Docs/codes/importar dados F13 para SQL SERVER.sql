


USE [DDS_EPOS_PRD]
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================================================================================
-- Proc Name   Load_F13_DAM
-- Description 
--  @PIF_File_Name: F13 DAM_EPOS_File Name (Excel format)
-- History
-- When        Who            What
-- 2017/04/07  William Wada      Initial version
-- =============================================================================================================
ALTER PROCEDURE [dbo].[Load_F13_DAM] 
@PIF_File_Name nvarchar(50) -- EXCEL format
AS
BEGIN


  SET NOCOUNT ON;

  declare @currDate    date;
  declare @dataCreatedDate    varchar(8);
  declare @sql         nvarchar(2048);
  declare @sheetname   nvarchar(50);

--Nome do arquivo ( Na verdade vai como parametro)
-- declare @PIF_File_Name nvarchar(50);
 --set @PIF_File_Name ='AF13_EPOS_WEB_SALES_20170328.xls'

  set @currDate = CONVERT(date,getdate(),111);
  set @dataCreatedDate = SUBSTRING(@PIF_File_Name,14,8) --Get from File Name
  set @sheetname = N'F13';


--
--LOAD EXCEL -> STG TABLE
--

  TRUNCATE TABLE dbo.STG_DAM_EPOS_F13;

  set @sql = 'insert into DDS_EPOS_PRD.dbo.STG_DAM_EPOS_F13'+
			N' SELECT NULL AS TYPE,'+
			N'RTRIM([�`���l���敪]) AS CHANNEL,'+
			N'RTRIM([�X�ܖ�]) AS STORE_NAME,'+
			N'RTRIM([�]�ƈ��J�i����]) AS EMPLOYEE_NAME,'+
			N'RTRIM([�X�^�b�t���]) AS STAFF_TYPE,'+
			N'RTRIM([���X���@]) AS MOT_VISIT,'+
			N'RTRIM([������]) AS DATE_SUBSCRIPTION,'+
			N'RTRIM([��������]) AS SUBSCRIPTION_TIME,'+
			N'RTRIM([�n����]) AS START_DATE,'+
			N'RTRIM([���i�敪]) AS CATEGORIES,'+
			N'RTRIM([�����v����]) AS SUBSCRIPTION_PLAN,'+
			N'RTRIM([�����R�[�X]) AS SUBSCRIPTION_COURSE,'+
			N'RTRIM([���z�ی���]) AS MP,'+
			N'RTRIM([��ی��Ґ���]) AS INSURED_GENDER,'+
			N'RTRIM([��ی��ҔN��]) AS INSURED_AGE,'+
			N'RTRIM([�J�[�h���]) AS CARD_TYPE,'+
			N'RTRIM([�E�H�[�����[�h�t���O]) AS WARM_LEAD_FLG,'+
			N'RTRIM([����A�b�v�Z���t���O]) AS CAN_UPSELL_FLG,'+
			N'RTRIM([CUSTOMER_ID]) AS CUSTOMER_ID,'+
			N'RTRIM([����E�z��҃t���O]) AS SPOUSE_FLG'+
			' FROM OPENROWSET(''Microsoft.ACE.OLEDB.12.0'',''Excel 8.0;HDR=YES;Database=\\172.31.20.38\ad_epos_prd\DATA_AREA\IN\DAM_F13\' +@PIF_File_Name + ''','''+
	' SELECT * FROM ['+ @sheetname +'$]'')';

  print @sql

  execute sp_executesql @sql



--
--TMP -> Final Table
--

--DECLARE @LAST_DATE_INSERTED varchar(25)

--set @LAST_DATE_INSERTED =select MAX(DATE_SUBSCRIPTION) from DAM_EPOS_F13

--INSERT INTO DAM_EPOS_F13
--select * 
--from STG_DAM_EPOS_F13
--where DATE_SUBSCRIPTION =@LAST_DATE_INSERTED
--;


END