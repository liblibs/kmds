CREATE TABLE [zInvoHead] (
  [ID] LONG,
  [CashSysID] LONG,
  [CashierID] SHORT,
  [ShiftID] LONG,
  [SalesTotal] CURRENCY,
  [TenderTotal] CURRENCY,
  [ClubInvoNo] LONG,
  [CheckNumber] VARCHAR (20),
  [Drawer] SHORT,
  [CheckingNumber] LONG,
  [CustomersODL] VARCHAR (50),
  [CustomersDOB] DATETIME,
  [SetupDate] DATETIME,
  [upsize_ts] BINARY,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([ID], [CashSysID])
)
