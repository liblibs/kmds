CREATE TABLE [zSalesPictures] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Text] VARCHAR (255),
  [FileName] VARCHAR (255) CONSTRAINT [IX_SalesPictures] UNIQUE,
  [SpecificItemOnly] BIT,
  [SetupDate] DATETIME,
  [ChangeDate] DATETIME,
  [upsize_ts] BINARY
)
