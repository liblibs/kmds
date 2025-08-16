SELECT
  InvoLine.CashSysID,
  InvoLine.InvoHeadID,
  InvoLine.SetupDate,
  First(InvoLine.Code) AS TenderCode
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.Price)<> 0
    )
    AND (
      (InvoLine.TransType)= "T"
    )
  )
GROUP BY
  InvoLine.CashSysID,
  InvoLine.InvoHeadID,
  InvoLine.SetupDate
HAVING
  (
    (
      (InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
    And (
      (
        First(InvoLine.Code)
      )< 100
    )
  );
