SELECT
  InvoLine.InvoHeadID,
  InvoLine.CashSysID,
  Max(InvoLine.SetupDate) AS MaxOfSetupDate,
  Max(InvoLine.SalesDept) AS MaxOfSalesDept INTO InvoLineClubInvoices
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.TransType)= "C"
    )
  )
GROUP BY
  InvoLine.InvoHeadID,
  InvoLine.CashSysID
HAVING
  (
    (
      (
        Max(InvoLine.SetupDate)
      )> Now()-400
    )
  );
