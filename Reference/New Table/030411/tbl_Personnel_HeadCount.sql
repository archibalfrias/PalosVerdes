if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Personnel_HeadCount]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_Personnel_HeadCount]
GO

CREATE TABLE [dbo].[tbl_Personnel_HeadCount] (
	[PK] [int] IDENTITY (1, 1) NOT NULL ,
	[LogInName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DivKey] [int] NOT NULL ,
	[DivName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DeptKey] [int] NOT NULL ,
	[DeptName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[StatusKey] [int] NOT NULL ,
	[StatusName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EmpCount] [float] NOT NULL 
) ON [PRIMARY]
GO

