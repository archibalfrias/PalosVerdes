if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbl_Inv_Items_tbl_Inv_Supplier]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbl_Inv_Items] DROP CONSTRAINT FK_tbl_Inv_Items_tbl_Inv_Supplier
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Inv_Supplier]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_Inv_Supplier]
GO

CREATE TABLE [dbo].[tbl_Inv_Supplier] (
	[PK] [int] NOT NULL ,
	[SupplierCode] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SupplierName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Type] [int] NOT NULL ,
	[Address1] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Address2] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Address3] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TelNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[FaxNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Email] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ContactPerson] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LastModified] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

