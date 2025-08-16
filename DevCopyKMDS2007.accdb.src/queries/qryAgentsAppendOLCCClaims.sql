INSERT INTO agents (
  AgentNumber, LastName, ActiveForTransfer
)
SELECT
  1000 AS AgentNumber,
  "OLCC Claims" AS FirstName,
  True AS Active;
