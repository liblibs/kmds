CREATE TABLE [OLCCWebItems] (
  [Code] LONG CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Description] VARCHAR (255),
  [Category] VARCHAR (50),
  [UCase] DOUBLE,
  [UOrder] LONG,
  [SetupDate] DATETIME,
  [SubCategoryID] AUTOINCREMENT
)
