SELECT
  InvoLine5.InvoHeadID,
  InvoLine5.ID,
  InvoLine5.CashSysID,
  InvoLine5.SalesDept AS ClubID,
  InvoLine5.SetupDate,
  Club.Name,
  Club.Licensee,
  Club.FintechAccount,
  CStr([SalesGroup]) AS Invoice_Number
FROM
  (
    qryFinTechTransactions
    INNER JOIN InvoLine5 ON (
      qryFinTechTransactions.SetupDate = InvoLine5.SetupDate
    )
    AND (
      qryFinTechTransactions.CashSysID = InvoLine5.CashSysID
    )
    AND (
      qryFinTechTransactions.InvoHeadID = InvoLine5.InvoHeadID
    )
  )
  INNER JOIN Club ON InvoLine5.SalesDept = Club.ID
WHERE
  (
    (
      (InvoLine5.ID)= 1
    )
  );
