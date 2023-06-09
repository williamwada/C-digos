USE [BTMU_PRD_DDS]
GO
/****** Object:  StoredProcedure [dbo].[Load_F8_Policy_data]    Script Date: 04/05/2017 18:15:35 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================================================================================
-- Proc Name   Load_F8_Policy_data
-- Description 
-- Parameters
--  @PIF_File_Name: F8 File Name (Excel format)
-- History
-- When        Who            What
-- 2017/04/04  Pham Lanh      Initial version
-- =============================================================================================================
ALTER PROCEDURE [dbo].[Load_F8_Policy_data] 
  @PIF_File_Name nvarchar(50) -- EXCEL format
AS
BEGIN
  -- SET NOCOUNT ON added to prevent extra result sets from
  -- interfering with SELECT statements.
  SET NOCOUNT ON;

  declare @success     int;
  declare @failure     int;
  declare @currDate    date;
  declare @dataCreatedDate    varchar(8);
  declare @Cancel_YM   varchar(8);
  declare @sql         nvarchar(2048);
  declare @sheetname   nvarchar(50);

  -- Set constant values
  set @success = 0;
  set @failure = 1;
  set @currDate = CONVERT(date,getdate(),111);
  set @dataCreatedDate = SUBSTRING(@PIF_File_Name,14,8) --Get from File Name
  set @Cancel_YM = CONVERT(VARCHAR(6),DATEADD(M,-1,CONVERT(date,@dataCreatedDate,111)),112) + '01' --previous created date
  set @sheetname = N'cÌ_ńŸŚĂȚ°À';

  TRUNCATE TABLE BTMU_PRD_STG.dbo.STG_IN_F8;

  set @sql = 'insert into BTMU_PRD_STG.dbo.STG_IN_F8'+
			N' SELECT RTRIM([ìŹPÊ]) AS CREATE_UNIT,'+
			N'RTRIM([ŰÔ]) AS POLICY_NUMBER,'+
			N'RTRIM([ÛŻnú]) AS POLICY_START_DATE,'+
			N'RTRIM([ÛŻIú]) AS POLICY_END_DATE,'+
			N'RTRIM([_ńÒŒ]) AS COMPANY_NAME,'+
			N'RTRIM([Áüú]) AS POLICY_ISSUE_DATE,'+
			N'RTRIM([tÔ]) AS COVERLETER_NUMBER,'+
			N'RTRIM([ÁüÒÔ]) AS POLICY_HOLDER_NUMBER,'+
			N'RTRIM([ÁüÒ«Ê]) AS GENDER,'+
			N'RTRIM([ÁüÒ¶Nú]) AS DOB,'+
			N'RTRIM([ÁüvÛŻż]) AS TOTAL_MP,'+
			N'RTRIM([ŸŚźÔP]) AS INFO1,'+
			N'RTRIM([ŸŚźÔQ]) AS POLICY_HOLDER_ID,'+
			N'RTRIM([ŸŚźÔR]) AS POLICY_ISSUE_DATE2,'+
			N'RTRIM([ŸŚźÔS]) AS CARD_COMPANY,'+
			N'RTRIM([ŸŚźÔT]) AS CAMPAIGN_NUMBER,'+
			N'RTRIM([ŸŚźÔU]) AS HAS_SPOUSE,'+
			N'RTRIM([ŸŚźÔV]) AS INFO7,'+
			N'RTRIM([ÁüÒÇÔ]) AS POLICY_MGT_NO,'+
			N'RTRIM([íÛŻÒÔ]) AS INSURED_NUMBER,'+
			N'RTRIM([{l\Š]) AS INSURED_MAIN,'+
			N'RTRIM([±ż]) AS INSURED_ID,'+
			N'RTRIM([íÛ@«Ê]) AS INSURED_GENDER,'+
			N'RTRIM([íÛ@¶Nú]) AS INSURED_DOB,'+
			N'RTRIM([íÛPńÛŻż]) AS INSURED_MP,'+
			N'RTRIM([EÆEíŒ]) AS OCCUPATION,'+
			N'RTRIM([NxŰÔ]) AS FIRST_POLICY_NUMBER,'+
			N'RTRIM([^]) AS TYPE,'+
			N'RTRIM([û]) AS KOUSU,'+
			N'RTRIM([^PńȘÛŻż]) AS TYPE_MP'+
	' FROM OPENROWSET(''Microsoft.ACE.OLEDB.12.0'',''Excel 8.0;HDR=YES;Database=D:\BTMUNK_PRD\STAGING_AREA\IN\F8\' +@PIF_File_Name + ''','''+
	' SELECT * FROM ['+ @sheetname +'$]'')';

  print @sql

  execute sp_executesql @sql

  print convert(varchar,@@rowcount)+' records loaded.';
  print ''


  print '--------------- UPDATE POLICY ---------------'

	UPDATE dbo.B_POLICY_DATA
	   SET CREATE_UNIT = F8.CREATE_UNIT
		  ,POLICY_NUMBER = F8.POLICY_NUMBER
		  ,POLICY_START_DATE = F8.POLICY_START_DATE
		  ,POLICY_END_DATE = F8.POLICY_END_DATE
		  ,COMPANY_NAME = F8.COMPANY_NAME
		  ,POLICY_ISSUE_DATE = F8.POLICY_ISSUE_DATE
		  ,COVERLETER_NUMBER = F8.COVERLETER_NUMBER
		  ,POLICY_HOLDER_NUMBER = F8.POLICY_HOLDER_NUMBER
		  ,GENDER = F8.GENDER
		  ,DOB = F8.DOB
		  ,TOTAL_MP = F8.TOTAL_MP
		  ,INFO1 = F8.INFO1
		  ,POLICY_HOLDER_ID = F8.POLICY_HOLDER_ID
		  ,POLICY_ISSUE_DATE2 = F8.POLICY_ISSUE_DATE2
		  ,CARD_COMPANY = F8.CARD_COMPANY
		  ,CAMPAIGN_NUMBER = F8.CAMPAIGN_NUMBER
		  ,HAS_SPOUSE = F8.HAS_SPOUSE
		  ,INFO7 = F8.INFO7
		  ,POLICY_MGT_NO = F8.POLICY_MGT_NO
		  ,INSURED_NUMBER = F8.INSURED_NUMBER
		  ,INSURED_MAIN = F8.INSURED_MAIN
		  ,INSURED_ID = F8.INSURED_ID
		  ,INSURED_GENDER = F8.INSURED_GENDER
		  ,INSURED_DOB = F8.INSURED_DOB
		  ,INSURED_MP = F8.INSURED_MP
		  ,OCCUPATION = F8.OCCUPATION
		  ,FIRST_POLICY_NUMBER = F8.FIRST_POLICY_NUMBER
		  ,TYPE = F8.TYPE
		  ,KOUSU = F8.KOUSU
		  ,TYPE_MP = F8.TYPE_MP
		  --,CANCEL_FLAG = F8.CANCEL_FLAG
		  --,CANCEL_YM = F8.CANCEL_YM
		  --,NEWEST_FLAG = 1
		  --,CREATE_DATETIME = F8.CREATE_DATETIME
		  ,UPDATE_DATETIME = GETDATE()
	FROM BTMU_PRD_STG.dbo.STG_IN_F8 F8
	WHERE B_POLICY_DATA.POLICY_HOLDER_ID = F8.POLICY_HOLDER_ID 
	AND CANCEL_FLAG = 0


  print convert(varchar,@@rowcount)+' records updated.';
  print ''

  UPDATE B_POLICY_DATA
  SET CANCEL_FLAG = 1
	, CANCEL_YM = @Cancel_YM
	, NEWEST_FLAG = 0
	, UPDATE_DATETIME = GETDATE()
  FROM B_POLICY_DATA PO
  WHERE not exists (select * from BTMU_PRD_STG.dbo.STG_IN_F8 F8 WHERE PO.POLICY_HOLDER_ID = F8.POLICY_HOLDER_ID)
  AND CANCEL_FLAG = 0
  
  print convert(varchar,@@rowcount)+' cancelled records updated.';
  print ''

  print '--------------- INSERT POLICY ---------------'
	INSERT INTO  dbo.B_POLICY_DATA
	SELECT CREATE_UNIT
		  ,POLICY_NUMBER
		  ,POLICY_START_DATE
		  ,POLICY_END_DATE
		  ,COMPANY_NAME
		  ,POLICY_ISSUE_DATE
		  ,COVERLETER_NUMBER
		  ,POLICY_HOLDER_NUMBER
		  ,GENDER
		  ,DOB
		  ,TOTAL_MP
		  ,INFO1
		  ,POLICY_HOLDER_ID
		  ,POLICY_ISSUE_DATE2
		  ,CARD_COMPANY
		  ,CAMPAIGN_NUMBER
		  ,HAS_SPOUSE
		  ,INFO7
		  ,POLICY_MGT_NO
		  ,INSURED_NUMBER
		  ,INSURED_MAIN
		  ,INSURED_ID
		  ,INSURED_GENDER
		  ,INSURED_DOB
		  ,INSURED_MP
		  ,OCCUPATION
		  ,FIRST_POLICY_NUMBER
		  ,TYPE
		  ,KOUSU
		  ,TYPE_MP
		  ,0 CANCEL_FLAG
		  ,NULL CANCEL_YM
		  ,1 NEWEST_FLAG
		  ,GETDATE() CREATE_DATETIME
		  ,GETDATE() UPDATE_DATETIME
	  FROM BTMU_PRD_STG.dbo.STG_IN_F8 F8
	  WHERE NOT EXISTS (SELECT * FROM B_POLICY_DATA PO WHERE PO.POLICY_HOLDER_ID = F8.POLICY_HOLDER_ID)

  print convert(varchar,@@rowcount)+' records Inserted.';
  print ''

	
END

