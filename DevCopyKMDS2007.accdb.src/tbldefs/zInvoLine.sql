CREATE TABLE [zInvoLine] (
  [InvoHeadID] LONG,
  [ID] LONG,
  [CashSysID] LONG,
  [Code] LONG,
  [SalesDept] LONG,
  [Quantity] LONG,
  [Price] CURRENCY,
  [InputCode] VARCHAR (15),
  [SalesGroup] LONG,
  [Cashier] LONG,
  [TransType] VARCHAR (1),
  [UpdateOnhand] BIT,
  [SetupDate] DATETIME,
  [upsize_ts] BINARY,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([InvoHeadID], [ID], [CashSysID])
)
