if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Member_Dependent]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_Member_Dependent]
GO

CREATE TABLE [dbo].[tbl_Member_Dependent] (
	[MemberKey] [int] NOT NULL ,
	[Line] [int] NOT NULL ,
	[ChildLName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ChildGName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ChildMName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ChildBirthDate] [datetime] NOT NULL ,
	[ChildPicture] [image] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

