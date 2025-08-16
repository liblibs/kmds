CREATE TABLE [LocLicensee] (
  [PremisesNumberImported] VARCHAR (255) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [TradeName] VARCHAR (60),
  [Street] VARCHAR (35),
  [City] VARCHAR (35),
  [Zip] VARCHAR (9),
  [PremisesNumber] LONG
)
