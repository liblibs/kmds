CREATE TABLE [InvoLineAllRegisters] (
  [InvoHeadID] LONG,
  [ID] LONG,
  [CashSysID] LONG,
  [CashierID] LONG,
  [ShiftID] LONG,
  [Code] VARCHAR,
  [SalesDept] LONG,
  [Quantity] LONG,
  [Price] CURRENCY,
  [InputCode] VARCHAR (15),
  [SalesGroup] LONG,
  [Drawer] SHORT,
  [TransType] VARCHAR (1),
  [UpdateOnhand] BIT,
  [SetupDate] DATETIME,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([InvoHeadID], [ID], [CashSysID])
)
