CREATE TABLE [NabcaOlccItem] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] PRIMARY KEY UNIQUE NOT NULL,
  [Code] VARCHAR,
  [OldCode] VARCHAR (6),
  [Description] VARCHAR (26),
  [Age] VARCHAR (4),
  [Proof] DOUBLE,
  [SizeDescription] VARCHAR (9),
  [Price] DOUBLE,
  [UnitsPerCase] LONG,
  [StockStatus] VARCHAR (1),
  [ItemID] LONG
)
