SELECT
  AgentTransferFTPImport.AgentNumber,
  AgentTransferFTPImport.Order,
  AgentTransferFTPImport.OrdDate
FROM
  AgentTransferFTPImport
GROUP BY
  AgentTransferFTPImport.AgentNumber,
  AgentTransferFTPImport.Order,
  AgentTransferFTPImport.OrdDate;
