CREATE TABLE [zSalesGroup] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] UNIQUE CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Description] VARCHAR (25) CONSTRAINT [Description] UNIQUE,
  [GroupLevel] LONG,
  [SetupDate] DATETIME,
  [ChangeDate] DATETIME,
  [upsize_ts] BINARY
)
