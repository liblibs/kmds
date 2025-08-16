CREATE TABLE [InventoryReportOrderBys] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ReportName] VARCHAR (50),
  [OrderByText] VARCHAR (50),
  [Ascending] BIT
)
