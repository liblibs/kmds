CREATE TABLE [zItemSalePrice] (
  [ItemID] VARCHAR,
  [StartDate] DATETIME,
  [EndDate] DATETIME,
  [OldPrice] CURRENCY,
  [SalePrice] CURRENCY,
  [SetupDate] DATETIME,
  [ChangeDate] DATETIME,
  [upsize_ts] BINARY,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([ItemID], [StartDate])
)
