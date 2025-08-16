SELECT
  AgentTransferHead.ID,
  AgentTransferHead.AgentID,
  AgentTransferHead.Order,
  AgentTransferHead.OrdDate,
  Agents.LastName AS AgentName,
  Agents.Address,
  [City] & ", " & [ST] & " " & [Zip] AS CitySTZ,
  Agents.Phone,
  AgentTransferLines.LineID,
  AgentTransferLines.ItemID,
  Item.Description,
  Item.SizeID,
  AgentTransferLines.QuantityShipped,
  AgentTransferLines.Price,
  qryUPCItemMax.UPCCodeMax,
  [QuantityShipped] * AgentTransferLines.Price AS ExtPrice,
  AgentTransferHead.Status,
  IIf(
    IsNull(Size.Description),
    "N/A",
    Size!Description
  ) AS SizeDesc,
  Item.Cost,
  [QuantityShipped] * Item.Cost AS ExtCost,
  AgentTransferLines.ClaimCode
FROM
  (
    (
      Agents
      INNER JOIN AgentTransferHead ON Agents.AgentNumber = AgentTransferHead.AgentID
    )
    INNER JOIN (
      (
        Item
        INNER JOIN AgentTransferLines ON Item.ID = AgentTransferLines.ItemID
      )
      LEFT JOIN qryUPCItemMax ON Item.ID = qryUPCItemMax.ItemID
    ) ON AgentTransferHead.ID = AgentTransferLines.AgentTransferHeadID
  )
  LEFT JOIN [Size] ON Item.SizeID = Size.ID
WHERE
  (
    (
      (AgentTransferHead.ID)= [Forms]![AgentsTransfersChoices]![txtAgentTransferID]
    )
  );
