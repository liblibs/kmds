CREATE TABLE [OLCCEPS] (
  [ID] AUTOINCREMENT,
  [AgencyNumber] LONG,
  [OrderNumber] LONG,
  [OLCCAlpha] VARCHAR (6),
  [OLCCItem] DOUBLE,
  [Description] VARCHAR (26),
  [UnitsPerCase] LONG,
  [Price] DOUBLE,
  [Ordered] LONG,
  [Shipped] LONG,
  [OrderDate] DATETIME,
  [SetupDate] DATETIME
)
