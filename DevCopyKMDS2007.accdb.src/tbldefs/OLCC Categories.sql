CREATE TABLE [OLCC Categories] (
  [Category] VARCHAR (255),
  [Code] DOUBLE CONSTRAINT [Code] PRIMARY KEY UNIQUE NOT NULL,
  [Units Per Case] DOUBLE,
  [Size] DOUBLE,
  [Description] VARCHAR (255),
  [UnitsPerOrder] LONG
)
