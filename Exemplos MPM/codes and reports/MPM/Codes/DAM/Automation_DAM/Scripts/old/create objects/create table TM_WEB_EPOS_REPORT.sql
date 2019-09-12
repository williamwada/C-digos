USE [EPOS]
GO

/****** Object:  Table [dbo].[TM_WebEposReport]    Script Date: 04/10/2017 16:34:15 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[TM_WebEposReport](
	[sales_date] [date] NULL,
	[Sales_Count_MGI] [int] NOT NULL,
	[Sales_Count_CAN] [int] NOT NULL,
	[MP_MGI] [int] NOT NULL,
	[MP_CAN] [int] NOT NULL
) ON [PRIMARY]

GO


