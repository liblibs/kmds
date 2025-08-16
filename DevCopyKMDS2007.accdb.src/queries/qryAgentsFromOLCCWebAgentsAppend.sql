INSERT INTO AGENTS (
  AgentNumber, LastName, Phone, Address,
  City, Zip
)
SELECT
  OLCCWebAgents.ID,
  OLCCWebAgents.Name,
  OLCCWebAgents.Phone,
  OLCCWebAgents.Address,
  OLCCWebAgents.Name,
  OLCCWebAgents.ZIP
FROM
  OLCCWebAgents;
