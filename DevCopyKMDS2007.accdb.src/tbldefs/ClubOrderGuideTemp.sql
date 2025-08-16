CREATE TABLE [ClubOrderGuideTemp] (
  [ClubID] LONG,
  [OrderID] LONG,
  [LineID] LONG,
  [LiquorItem] BYTE,
  [ItemID] LONG,
  [Description] VARCHAR (25),
  [Par] LONG,
  [QuotedPrice] CURRENCY,
  [SetupDate] DATETIME,
   CONSTRAINT ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([ClubID], [OrderID], [LineID], [LiquorItem])
)
