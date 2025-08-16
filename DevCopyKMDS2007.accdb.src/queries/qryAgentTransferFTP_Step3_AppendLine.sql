INSERT INTO AgentTransferLines (
  LineID, ItemID, QuantityOrdered, QuantityShipped,
  Price, Cost, AgentTransferHeadID
)
SELECT
  AgentTransferFTPImport.LineID,
  AgentTransferFTPImport.ItemID,
  AgentTransferFTPImport.QuantityShipped,
  AgentTransferFTPImport.QuantityShipped,
  AgentTransferFTPImport.Price,
  AgentTransferFTPImport.Cost,
  qryAgentTransferHeadIDFromFTP.ID
FROM
  qryAgentTransferHeadIDFromFTP,
  AgentTransferFTPImport
  INNER JOIN Item ON AgentTransferFTPImport.ItemID = Item.ID;
