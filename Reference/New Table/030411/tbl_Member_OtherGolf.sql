if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Member_OtherGolf]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_Member_OtherGolf]
GO

CREATE TABLE [dbo].[tbl_Member_OtherGolf] (
	[MemberKey] [int] NOT NULL ,
	[Line] [int] NOT NULL ,
	[OtherGolfClubs] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MemberSince] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

