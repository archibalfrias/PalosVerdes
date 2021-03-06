if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbl_Inv_Items_tbl_Inv_Class]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbl_Inv_Items] DROP CONSTRAINT FK_tbl_Inv_Items_tbl_Inv_Class
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Inv_Class]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_Inv_Class]
GO

CREATE TABLE [dbo].[tbl_Inv_Class] (
	[PK] [int] NOT NULL ,
	[ClassCode] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectKey] [int] NOT NULL ,
	[LastModified] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

