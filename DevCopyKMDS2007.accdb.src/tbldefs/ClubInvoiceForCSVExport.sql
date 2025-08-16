CREATE TABLE [ClubInvoiceForCSVExport] (
  [AgentNumber] LONG,
  [Name] VARCHAR (40),
  [ClubID] LONG,
  [ClubName] VARCHAR (50),
  [InvoiceNo] LONG,
  [InvDate] DATETIME,
  [ItemID] LONG,
  [Description] VARCHAR (50),
  [Size] VARCHAR (25),
  [Quantity] LONG,
  [Price] CURRENCY,
  [XPrice] CURRENCY,
  [SalePrice] VARCHAR (1),
  [PricePerOZ] CURRENCY
)
