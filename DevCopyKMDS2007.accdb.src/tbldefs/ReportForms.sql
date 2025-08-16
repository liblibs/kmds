CREATE TABLE [ReportForms] (
  [ReportName] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ReportType] VARCHAR (50),
  [SetupDate] DATETIME,
  [ChangeDate] DATETIME
)
