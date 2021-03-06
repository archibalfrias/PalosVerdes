if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbl_Inv_Supplier_tbl_Inv_SupplierType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbl_Inv_Supplier] DROP CONSTRAINT FK_tbl_Inv_Supplier_tbl_Inv_SupplierType
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Inv_SupplierType]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_Inv_SupplierType]
GO

CREATE TABLE [dbo].[tbl_Inv_SupplierType] (
	[PK] [int] NOT NULL ,
	[SupplierType] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

