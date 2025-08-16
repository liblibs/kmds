CREATE TABLE [WebOrders] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] PRIMARY KEY UNIQUE NOT NULL,
  [WebOrderId] VARCHAR (255),
  [ClubID] LONG,
  [LineId] LONG,
  [ItemID] LONG,
  [Quantity] LONG,
  [ExtendNet] DOUBLE,
  [Net] DOUBLE,
  [Retail] DOUBLE
)
