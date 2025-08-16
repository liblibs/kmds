CREATE TABLE [zSalesLeads] (
  [id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Text] VARCHAR (255),
  [SpecificItemOnly] BIT,
  [SetupDate] DATETIME,
  [ChangeDate] DATETIME,
  [upsize_ts] BINARY
)
