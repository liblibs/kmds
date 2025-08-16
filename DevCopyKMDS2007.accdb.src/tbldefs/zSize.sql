CREATE TABLE [zSize] (
  [ID] LONG CONSTRAINT [ID] UNIQUE CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Description] VARCHAR (25) CONSTRAINT [Description] UNIQUE,
  [Oz] DOUBLE,
  [CaseQty] LONG,
  [SetupDate] DATETIME,
  [ChangeDate] DATETIME,
  [upsize_ts] BINARY
)
