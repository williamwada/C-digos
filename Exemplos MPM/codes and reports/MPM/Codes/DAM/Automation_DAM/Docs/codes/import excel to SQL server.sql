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
  set @sheetname = N'ícëÃå_ñÒñæç◊√ﬁ∞¿';

  TRUNCATE TABLE BTMU_PRD_STG.dbo.STG_IN_F8;

  set @sql = 'insert into BTMU_PRD_STG.dbo.STG_IN_F8'+
			N' SELECT RTRIM([çÏê¨íPà ]) AS CREATE_UNIT,'+
			N'RTRIM([èÿåîî‘çÜ]) AS POLICY_NUMBER,'+
			N'RTRIM([ï€åØénä˙]) AS POLICY_START_DATE,'+
			N'RTRIM([ï€åØèIä˙]) AS POLICY_END_DATE,'+
			N'RTRIM([å_ñÒé“ñº]) AS COMPANY_NAME,'+
			N'RTRIM([â¡ì¸ì˙]) AS POLICY_ISSUE_DATE,'+
			N'RTRIM([ëóïtèëî‘çÜ]) AS COVERLETER_NUMBER,'+
			N'RTRIM([â¡ì¸é“î‘çÜ]) AS POLICY_HOLDER_NUMBER,'+
			N'RTRIM([â¡ì¸é“ê´ï ]) AS GENDER,'+
			N'RTRIM([â¡ì¸é“ê∂îNåéì˙]) AS DOB,'+
			N'RTRIM([â¡ì¸çáåvï€åØóø]) AS TOTAL_MP,'+
			N'RTRIM([ñæç◊êÆóùî‘çÜÇP]) AS INFO1,'+
			N'RTRIM([ñæç◊êÆóùî‘çÜÇQ]) AS POLICY_HOLDER_ID,'+
			N'RTRIM([ñæç◊êÆóùî‘çÜÇR]) AS POLICY_ISSUE_DATE2,'+
			N'RTRIM([ñæç◊êÆóùî‘çÜÇS]) AS CARD_COMPANY,'+
			N'RTRIM([ñæç◊êÆóùî‘çÜÇT]) AS CAMPAIGN_NUMBER,'+
			N'RTRIM([ñæç◊êÆóùî‘çÜÇU]) AS HAS_SPOUSE,'+
			N'RTRIM([ñæç◊êÆóùî‘çÜÇV]) AS INFO7,'+
			N'RTRIM([â¡ì¸é“ä«óùî‘çÜ]) AS POLICY_MGT_NO,'+
			N'RTRIM([îÌï€åØé“î‘çÜ]) AS INSURED_NUMBER,'+
			N'RTRIM([ñ{êlï\é¶]) AS INSURED_MAIN,'+
			N'RTRIM([ë±ïø]) AS INSURED_ID,'+
			N'RTRIM([îÌï€Å@ê´ï ]) AS INSURED_GENDER,'+
			N'RTRIM([îÌï€Å@ê∂îNåéì˙]) AS INSURED_DOB,'+
			N'RTRIM([îÌï€ÇPâÒï€åØóø]) AS INSURED_MP,'+
			N'RTRIM([êEã∆êEéÌñº]) AS OCCUPATION,'+
			N'RTRIM([èâîNìxèÿåîî‘çÜ]) AS FIRST_POLICY_NUMBER,'+
			N'RTRIM([å^]) AS TYPE,'+
			N'RTRIM([å˚êî]) AS KOUSU,'+
			N'RTRIM([å^ÇPâÒï™ï€åØóø]) AS TYPE_MP'+
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

