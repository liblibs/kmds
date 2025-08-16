CREATE TABLE [zSalesDepartment] (
  [ID] LONG CONSTRAINT [ID] UNIQUE CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [SalesGroupID] LONG,
  [Description] VARCHAR (50) CONSTRAINT [Description] UNIQUE,
  [Discounts] BIT,
  [NegativeOnhand] BIT,
  [AllowDiscounts] BIT,
  [IncludeInInventoryValueReports] BIT,
  [TargetGP] DOUBLE,
  [AllowClubDiscounts] BIT,
  [CatchAllItemID] VARCHAR,
  [MinimumAgeToBuy] LONG,
  [SetupDate] DATETIME,
  [ChangeDate] DATETIME,
  [upsize_ts] BINARY
)
