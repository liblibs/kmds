INSERT INTO AgentTransferHead (
  AgentID, [Order], OrdDate, Status,
  OrderNotes
)
SELECT
  AgentTransferFTPImport.AgentNumber,
  AgentTransferFTPImport.Order,
  AgentTransferFTPImport.OrdDate,
  "I" AS Status,
  "Imported from KMDS FTP" AS OrderNotes
FROM
  AgentTransferFTPImport
GROUP BY
  AgentTransferFTPImport.AgentNumber,
  AgentTransferFTPImport.Order,
  AgentTransferFTPImport.OrdDate,
  "I",
  "Imported from KMDS FTP";
