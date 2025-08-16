CREATE TABLE [InvoiceLineLD] (
  [InvoHeadID] LONG,
  [ID] LONG,
  [CashSysID] LONG,
  [CashierID] LONG,
  [ShiftID] LONG,
  [Code] LONG,
  [SalesDept] LONG,
  [Quantity] LONG,
  [Price] CURRENCY,
  [InputCode] VARCHAR (255),
  [SalesGroup] LONG,
  [Drawer] SHORT,
  [TransType] VARCHAR (255),
  [UpdateOnhand] BIT,
  [SetupDate] DATETIME,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([InvoHeadID], [ID], [CashSysID], [SetupDate])
)
