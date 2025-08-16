CREATE TABLE [OLCCAgents] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] PRIMARY KEY UNIQUE NOT NULL,
  [AgentNumber] LONG,
  [AgencyName] VARCHAR (255),
  [Phone] VARCHAR (255),
  [District] DOUBLE,
  [Location] VARCHAR (255),
  [FirstName] VARCHAR (255),
  [LastName] VARCHAR (255),
  [Add1] VARCHAR (255),
  [Add2] VARCHAR (255),
  [City] VARCHAR (255),
  [ZipCode] VARCHAR (255),
  [Hours] VARCHAR (255)
)
