CREATE TABLE [zTender] (
  [ID] LONG CONSTRAINT [ID] UNIQUE CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Description] VARCHAR (25),
  [ChangeAllowed] BIT,
  [ManagerOnly] BIT,
  [GetCheckCashing] BIT,
  [GetCheckNumber] BIT,
  [NonCashDetailReport] BIT,
  [OpenCashDrawer] BIT,
  [SetupDate] DATETIME,
  [ChangeDate] DATETIME,
  [upsize_ts] BINARY
)
