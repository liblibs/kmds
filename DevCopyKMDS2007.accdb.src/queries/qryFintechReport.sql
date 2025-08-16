SELECT
  qryFintechDivisionID.Division_ID,
  qryFintechTransactionClubID.Invoice_Number,
  InvoLine5.SETUPDATE AS invoice_date,
  qryFintechTransactionClubID.Licensee AS Vendor_store_id,
  0 - [Quantity] AS quantity_shipped,
  "EA" AS Quantity_uom,
  [involine5].[price] / [involine5].[quantity] AS unit_price,
  Cashier.ID AS CashierID,
  Cashier.Name AS CashierName,
  InvoLine5.CashSysID,
  qryFintechTransactionClubID.Name AS ClubName,
  InvoLine5.TransType,
  0 AS amount
FROM
  Sysfile,
  qryFintechDivisionID,
  Cashier
  INNER JOIN (
    (
      qryFinTechTransactions
      INNER JOIN qryFintechTransactionClubID ON (
        qryFinTechTransactions.InvoHeadID = qryFintechTransactionClubID.InvoHeadID
      )
      AND (
        qryFinTechTransactions.CashSysID = qryFintechTransactionClubID.CashSysID
      )
      AND (
        qryFinTechTransactions.SetupDate = qryFintechTransactionClubID.SetupDate
      )
    )
    INNER JOIN InvoLine5 ON (
      qryFinTechTransactions.InvoHeadID = InvoLine5.InvoHeadID
    )
    AND (
      qryFinTechTransactions.CashSysID = InvoLine5.CashSysID
    )
    AND (
      qryFinTechTransactions.SetupDate = InvoLine5.SetupDate
    )
  ) ON Cashier.ID = InvoLine5.CashierID
WHERE
  (
    (
      (InvoLine5.TransType)= "T"
    )
  );
