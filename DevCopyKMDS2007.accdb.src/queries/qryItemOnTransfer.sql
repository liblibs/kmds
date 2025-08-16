SELECT
  AgentTransferHead.ID,
  AgentTransferHead.OrdDate,
  Agents.AgentNumber,
  Agents.FirstName,
  Agents.LastName,
  AgentTransferLines.ItemID,
  AgentTransferLines.QuantityOrdered,
  AgentTransferLines.Price,
  AgentTransferLines.Cost,
  AgentTransferLines.SetupDate,
  AgentTransferLines.ChangeDate,
  [QuantityShipped] * [Price] AS ExtPrice,
  [QuantityShipped] * [Cost] AS ExtCost,
  IIf([Status] = "I", "In", "Out") AS InOrOut
FROM
  (
    AgentTransferLines
    INNER JOIN AgentTransferHead ON AgentTransferLines.AgentTransferHeadID = AgentTransferHead.ID
  )
  INNER JOIN Agents ON AgentTransferHead.AgentID = Agents.AgentNumber
WHERE
  (
    (
      (
        AgentTransferLines.QuantityOrdered
      )<> 0
    )
  );
