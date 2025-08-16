PARAMETERS parAgentTransferHead Long;
UPDATE
  Sysfile,
  AgentTransferHead
SET
  AgentTransferHead.InvoiceNo = Sysfile.InvoiceCounter,
  AgentTransferHead.InvDate = Now()
WHERE
  (
    (
      (AgentTransferHead.ID)= [parAgentTransferHead]
    )
  );
