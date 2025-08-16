CREATE TABLE [zBuyDownItem] (
  [BuyDownID] LONG,
  [Line] LONG,
  [Club] LONG,
  [Order] LONG,
  [ItemsCode] DOUBLE CONSTRAINT [ItemsCode] UNIQUE,
  [QuantitySold] LONG,
  [DiscountDollars] CURRENCY,
  [ShpQty] LONG,
  [ShpDollars] CURRENCY,
  [SpcPrice] CURRENCY,
  [SetupDate] DATETIME,
  [ChangeDate] DATETIME,
  [LastSoldDate] DATETIME,
  [upsize_ts] BINARY,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([BuyDownID], [Line])
)
