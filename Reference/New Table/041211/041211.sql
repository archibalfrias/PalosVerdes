if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbl_GolfCart_Info_tbl_GolfCart_Owner_Type]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbl_GolfCart_Info] DROP CONSTRAINT FK_tbl_GolfCart_Info_tbl_GolfCart_Owner_Type
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Caddy_Information]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_Caddy_Information]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_GolfCart_Info]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_GolfCart_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_GolfCart_Owner_Type]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_GolfCart_Owner_Type]
GO

CREATE TABLE [dbo].[tbl_Caddy_Information] (
	[PK] [int] IDENTITY (1, 1) NOT NULL ,
	[CaddyNo] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CaddyLName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CaddyFName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CaddyMName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CaddyContactNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CaddyContactPerson] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CaddyContactPersonNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LastModified] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_GolfCart_Info] (
	[PK] [int] IDENTITY (1, 1) NOT NULL ,
	[GolfCartNo] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ChasisNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EngineNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[OwnerType] [int] NOT NULL ,
	[Owner] [int] NULL ,
	[CoOwner] [int] NULL ,
	[Description] [varchar] (50) NOT NULL ,
	[LastModified] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tbl_GolfCart_Owner_Type] (
	[PK] [int] IDENTITY (1, 1) NOT NULL ,
	[sName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

