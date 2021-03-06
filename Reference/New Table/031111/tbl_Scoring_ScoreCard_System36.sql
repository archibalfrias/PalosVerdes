if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbl_Scoring_ScoreCard_System36_Detail_tbl_Scoring_ScoreCard_System36]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbl_Scoring_ScoreCard_System36_Detail] DROP CONSTRAINT FK_tbl_Scoring_ScoreCard_System36_Detail_tbl_Scoring_ScoreCard_System36
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Scoring_ScoreCard_System36]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_Scoring_ScoreCard_System36]
GO

CREATE TABLE [dbo].[tbl_Scoring_ScoreCard_System36] (
	[PK] [int] IDENTITY (1, 1) NOT NULL ,
	[CtrlNo] [varchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TournamentKey] [int] NOT NULL ,
	[PlayerKey] [int] NOT NULL ,
	[DDate] [datetime] NOT NULL ,
	[Front9Gross] [float] NOT NULL ,
	[Back9Gross] [float] NOT NULL ,
	[GrossPoints] AS ([Front9Gross] + [Back9Gross]) ,
	[Front9Net] [float] NOT NULL ,
	[Back9Net] [float] NOT NULL ,
	[NetPoints] AS ([Front9Net] + [Back9Net]) ,
	[HDCP] [float] NOT NULL ,
	[Eagle] [float] NOT NULL ,
	[Birdie] [float] NOT NULL ,
	[Par] [float] NOT NULL ,
	[Boogie] [float] NOT NULL ,
	[Boogie_2] [float] NOT NULL ,
	[Boogie_3] [float] NOT NULL ,
	[LastModified] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

