if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Member_CardInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_Member_CardInfo]
GO

CREATE TABLE [dbo].[tbl_Member_CardInfo] (
	[MemberKey] [int] NOT NULL ,
	[Line] [int] NOT NULL ,
	[CardAccount] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CardType] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

