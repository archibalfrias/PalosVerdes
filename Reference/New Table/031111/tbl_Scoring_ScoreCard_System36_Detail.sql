if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Scoring_ScoreCard_System36_Detail]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_Scoring_ScoreCard_System36_Detail]
GO

CREATE TABLE [dbo].[tbl_Scoring_ScoreCard_System36_Detail] (
	[ScoreCardKey] [int] NOT NULL ,
	[Hole] [int] NOT NULL ,
	[Par] [float] NOT NULL ,
	[Score] [float] NOT NULL 
) ON [PRIMARY]
GO

