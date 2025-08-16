CREATE TABLE [OLCC_8_and_12_UPC] (
  [Code] LONG,
  [lid] LONG,
  [ThirdDigit] VARCHAR (255),
  [UPCCode] VARCHAR (14) CONSTRAINT [UPCCode] UNIQUE,
  [upc12] VARCHAR (255)
)
