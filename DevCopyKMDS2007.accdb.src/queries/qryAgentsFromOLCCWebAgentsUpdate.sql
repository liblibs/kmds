UPDATE
  OLCCWebAgents
  INNER JOIN Agents ON OLCCWebAgents.ID = Agents.AgentNumber
SET
  Agents.LastName = OLCCWebAgents.name,
  Agents.Address = OLCCWebAgents.address,
  Agents.City = OLCCWebAgents.City,
  Agents.ST = "OR",
  Agents.Zip = OLCCWebAgents.ZIP,
  Agents.Phone = OLCCWebAgents.PHONE;
