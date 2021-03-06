if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Personnel_Active_Inactive_Report]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_Personnel_Active_Inactive_Report]
GO

CREATE TABLE [dbo].[tbl_Personnel_Active_Inactive_Report] (
	[PK] [int] IDENTITY (1, 1) NOT NULL ,
	[LogInName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Division] [int] NOT NULL ,
	[DivisionName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Department] [int] NOT NULL ,
	[DepartmentName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[StatusKey] [int] NOT NULL ,
	[StatusName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PositionKey] [int] NOT NULL ,
	[PositionName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EmpKey] [int] NOT NULL ,
	[IDNumber] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EmployeeName] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EffecDate] [datetime] NULL ,
	[Reason] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

