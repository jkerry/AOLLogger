SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [LoggedItems](
	[Folder] [nvarchar](50) NULL,
	[Subject] [nvarchar](256) NULL,
	[Body] [nvarchar](max) NULL,
	[RecieptTime] [datetime] NULL,
	[id] [bigint] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO

