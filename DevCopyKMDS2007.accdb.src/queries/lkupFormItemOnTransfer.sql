SELECT
  AgentTransferLines.ItemID,
  Sum(
    AgentTransferLines.QuantityOrdered
  ) AS OnTransfer
FROM
  AgentTransferLines
GROUP BY
  AgentTransferLines.ItemID;
