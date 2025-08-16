SELECT
  AdjustReasons.Reason
FROM
  AdjustReasons
WHERE
  (
    (
      (AdjustReasons.Reason)= [Forms]![AgentsListBox]![AgentsActiveForTransferFalse].[Form]![ActiveForTransfer]
    )
  );
