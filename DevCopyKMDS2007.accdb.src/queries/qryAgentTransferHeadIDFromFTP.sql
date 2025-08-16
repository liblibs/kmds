SELECT
  AgentTransferHead.ID,
  AgentTransferHead.Status
FROM
  qryAgentTransferFromFTPHead
  INNER JOIN AgentTransferHead ON (
    qryAgentTransferFromFTPHead.OrdDate = AgentTransferHead.OrdDate
  )
  AND (
    qryAgentTransferFromFTPHead.Order = AgentTransferHead.Order
  )
  AND (
    qryAgentTransferFromFTPHead.AgentNumber = AgentTransferHead.AgentID
  )
WHERE
  (
    (
      (AgentTransferHead.Status)= "I"
    )
  );
