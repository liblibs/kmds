PARAMETERS parAgentNumber Long;
SELECT
  Agents.AgentNumber,
  Agents.OrderCount
FROM
  Agents
WHERE
  (
    (
      (Agents.AgentNumber)= [parAgentNumber]
    )
  );
