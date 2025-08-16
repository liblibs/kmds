SELECT
  UPCItem.SetupDate,
  UPCItem.ID,
  UPCItem.ItemID,
  qryInvoLineUPC_MRU.MaxOfSetupDate,
  Len([id]) AS ll INTO upcOld
FROM
  UPCItem
  LEFT JOIN qryInvoLineUPC_MRU ON UPCItem.ID = qryInvoLineUPC_MRU.InputCode
WHERE
  (
    (
      (UPCItem.SetupDate)< Now()-30
    )
    AND (
      (
        qryInvoLineUPC_MRU.MaxOfSetupDate
      )< Now()-180
    )
  )
  OR (
    (
      (UPCItem.SetupDate)< Now()-30
    )
    AND (
      (
        qryInvoLineUPC_MRU.MaxOfSetupDate
      ) Is Null
    )
    AND (
      (
        Len([id])
      )> 7
    )
  );
