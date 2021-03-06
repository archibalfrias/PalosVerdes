if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbl_Member_CardInfo_tbl_Member_Information]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbl_Member_CardInfo] DROP CONSTRAINT FK_tbl_Member_CardInfo_tbl_Member_Information
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbl_Member_Dependent_tbl_Member_Information]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbl_Member_Dependent] DROP CONSTRAINT FK_tbl_Member_Dependent_tbl_Member_Information
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tbl_Member_OtherGolf_tbl_Member_Information]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbl_Member_OtherGolf] DROP CONSTRAINT FK_tbl_Member_OtherGolf_tbl_Member_Information
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Member_Information]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbl_Member_Information]
GO

CREATE TABLE [dbo].[tbl_Member_Information] (
	[PK] [int] IDENTITY (1, 1) NOT NULL ,
	[LastName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[FirstName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MiddleName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Residence] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BirthPlace] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BirthDate] [datetime] NOT NULL ,
	[ContactNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EmailAdd] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Gender] [tinyint] NOT NULL ,
	[CivilStatus] [tinyint] NOT NULL ,
	[TIN] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Citizenship] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Citizenship1] [datetime] NULL ,
	[ResCertNo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ResCertNo1] [datetime] NULL ,
	[ResCertNo2] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CollegeUniversity] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DegreeObtained] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Affiliation] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MemberPicture] [image] NULL ,
	[SpouseLName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SpouseGName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SpouseMName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SpouseContact] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SpouseOccupation] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SpouseCompany] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SpouseCollege] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SpouseDegreeObtained] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SpousePicture] [image] NULL ,
	[BusinessName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BusinessPosition] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BusinessTel] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BusinessAddress] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BusinessFax] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BusinessNature] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LastModified] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

