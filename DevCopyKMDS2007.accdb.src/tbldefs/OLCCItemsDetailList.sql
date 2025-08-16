CREATE TABLE [OLCCItemsDetailList] (
  [ID] AUTOINCREMENT,
  [OLCCCode] VARCHAR (255) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Desription] VARCHAR (255),
  [Category] VARCHAR (255),
  [Size] VARCHAR (255),
  [Status] VARCHAR (255),
  [UnitsPerOrder] DOUBLE,
  [UnitsPerCase] DOUBLE,
  [PricePerBottle] VARCHAR (255),
  [QuantityAvailable] VARCHAR (255),
  [ItemID] LONG,
  [Field11] VARCHAR (255)
)
