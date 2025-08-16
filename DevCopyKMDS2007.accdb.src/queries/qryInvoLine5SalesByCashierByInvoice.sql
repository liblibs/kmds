SELECT
  InvoLine5.CashierID,
  InvoLine5.CashSysID,
  InvoLine5.InvoHeadID,
  Sum(
    IIf([TransType] = "S", 1, 0)
  ) AS TransactionLines,
  Sum(
    IIf([TransType] = "S", [Price], 0)
  ) AS Sales,
  Max(
    IIf([TransType] = "S", 1, 0)
  ) AS Transactions,
  Max(
    IIf(
      [TransType] = "T"
      And [CODE] = [NoSaleTenderID],
      1, 0
    )
  ) AS NoSales,
  Max(
    IIf(
      [TransType] = "T"
      And [CODE] = [CancelTenderID],
      1, 0
    )
  ) AS CancelledSales,
  Sum(
    IIf(
      [TransType] = "T"
      And [CODE] = [ScannedODLTenderID],
      1, 0
    )
  ) AS IDScanCount
FROM
  InvoLine5,
  Sysfile
GROUP BY
  InvoLine5.CashierID,
  InvoLine5.CashSysID,
  InvoLine5.InvoHeadID
ORDER BY
  InvoLine5.CashSysID,
  InvoLine5.InvoHeadID;
