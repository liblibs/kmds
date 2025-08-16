SELECT
  AgentTransferLines.ItemID,
  Sum(
    AgentTransferLines.QuantityOrdered
  ) AS SumOfQuantityOrdered
FROM
  AgentTransferLines
GROUP BY
  AgentTransferLines.ItemID;
