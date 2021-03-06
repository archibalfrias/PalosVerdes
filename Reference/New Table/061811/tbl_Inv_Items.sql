if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Inv_Items]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_Inv_Items]
GO

CREATE TABLE [dbo].[tbl_Inv_Items] (
	[PK] [int] NOT NULL ,
	[ItemCode] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ItemDesc] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SuppKey] [int] NOT NULL ,
	[ClassKey] [int] NOT NULL ,
	[Unit] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Unit2] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ConUnit] [float] NOT NULL ,
	[ConUnit2] [float] NOT NULL ,
	[Cost] [float] NOT NULL ,
	[SRP] [float] NOT NULL ,
	[MaxQty] [float] NOT NULL ,
	[MinQty] [float] NOT NULL ,
	[Remarks] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LastModified] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

