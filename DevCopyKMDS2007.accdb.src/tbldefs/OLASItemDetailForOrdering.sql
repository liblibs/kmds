CREATE TABLE [OLASItemDetailForOrdering] (
  [ID] AUTOINCREMENT,
  [ItemCode] VARCHAR (255),
  [NewItemCode] VARCHAR CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Description] VARCHAR (50),
  [Category] VARCHAR (50),
  [Size] VARCHAR (10),
  [Status] VARCHAR (50),
  [NoWillCalls] VARCHAR (1),
  [MinOrderQty] LONG,
  [MaxOrderQty] LONG,
  [BottlesPerCase] LONG,
  [PricePerBottle] CURRENCY,
  [QuantityAvailable] LONG,
  [ReportDate] DATETIME,
  [SetupDate] DATETIME
)
