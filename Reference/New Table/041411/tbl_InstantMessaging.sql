if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_InstantMessaging]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_InstantMessaging]
GO

CREATE TABLE [dbo].[tbl_InstantMessaging] (
	[PK] [int] IDENTITY (1, 1) NOT NULL ,
	[Date_Time] [datetime] NOT NULL ,
	[Message] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[From_User] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[To_User] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MsgType] [tinyint] NOT NULL ,
	[Opened] [tinyint] NOT NULL 
) ON [PRIMARY]
GO

