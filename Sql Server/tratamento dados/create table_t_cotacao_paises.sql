USE [teste_API]
GO

/****** Object:  Table [dbo].[t_cotacao_paises]    Script Date: 15/04/2019 10:45:55 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[t_cotacao_paises](
	[serie] [numeric](18, 0) NOT NULL,
	[valor] [numeric](18, 4) NULL,
	[data] [date] NULL,
	[pais] [nchar](50) NULL
) ON [PRIMARY]
GO


