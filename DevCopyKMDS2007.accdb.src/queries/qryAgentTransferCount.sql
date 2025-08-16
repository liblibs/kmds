PARAMETERS parAgentID Long;
SELECT
  Agents.AgentNumber,
  Agents.OrderCount
FROM
  Agents
WHERE
  (
    (
      (Agents.AgentNumber)= [parAgentID]
    )
  );
