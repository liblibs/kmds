SELECT
  Sysfile.AgentNumber,
  Sysfile.Name,
  KMDSKS.SetupDate
FROM
  Sysfile
  INNER JOIN KMDSKS ON Sysfile.AgentNumber = KMDSKS.AgentNumber;
