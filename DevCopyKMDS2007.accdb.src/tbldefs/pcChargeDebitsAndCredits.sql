CREATE TABLE [pcChargeDebitsAndCredits] (
  [Station] VARCHAR (8),
  [Date] DATETIME,
  [Time] VARCHAR (8),
  [Number] VARCHAR (10),
  [Card] VARCHAR (20),
  [Member] VARCHAR (20),
  [Result] VARCHAR (21),
  [Auth] VARCHAR (16),
  [Amount] CURRENCY,
  [CapturedAmount] CURRENCY,
  [DebitOrCredit] VARCHAR (1),
  [TicketNumber] LONG
)
