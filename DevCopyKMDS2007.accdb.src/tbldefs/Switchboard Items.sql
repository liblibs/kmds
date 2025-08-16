CREATE TABLE [Switchboard Items] (
  [SwitchboardID] LONG,
  [ItemNumber] SHORT,
  [ItemText] VARCHAR (255),
  [Command] SHORT,
  [Argument] VARCHAR (255),
  [upsize_ts] BINARY,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([SwitchboardID], [ItemNumber])
)
