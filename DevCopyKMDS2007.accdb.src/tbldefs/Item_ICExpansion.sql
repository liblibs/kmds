CREATE TABLE [Item_ICExpansion] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] PRIMARY KEY UNIQUE NOT NULL,
  [Code] DOUBLE,
  [Description] VARCHAR (26),
  [Age] VARCHAR (4),
  [Proof] DOUBLE,
  [SizeDesc] VARCHAR (9),
  [Price] CURRENCY,
  [UnitsPerCase] LONG,
  [StockStatus] VARCHAR (1),
  [OldItemID] LONG,
  [OldItemAlpha] VARCHAR (6)
)
