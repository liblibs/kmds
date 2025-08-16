SELECT
  AgentTransferFTPImport.ItemID,
  AgentTransferFTPImport.Description,
  Item.ID,
  AgentTransferFTPImport.LineID,
  AgentTransferFTPImport.ItemID,
  AgentTransferFTPImport.Description,
  AgentTransferFTPImport.SizeID,
  AgentTransferFTPImport.QuantityShipped,
  AgentTransferFTPImport.Price,
  AgentTransferFTPImport.UPCCodeMax,
  AgentTransferFTPImport.ExtPrice,
  AgentTransferFTPImport.Status,
  AgentTransferFTPImport.SizeDesc
FROM
  AgentTransferFTPImport
  LEFT JOIN Item ON AgentTransferFTPImport.ItemID = Item.ID
WHERE
  (
    (
      (Item.ID) Is Null
    )
  );
