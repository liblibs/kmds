UPDATE
  AgentTransferFTPImport
  INNER JOIN UPCItem ON AgentTransferFTPImport.UPCCodeMax = UPCItem.ID
SET
  AgentTransferFTPImport.ItemID = UPCItem.ItemID
WHERE
  (
    (
      (AgentTransferFTPImport.ItemID) Not Between 1000
      And 99999
    )
  );
