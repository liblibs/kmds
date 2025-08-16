SELECT
  InvoLine.SetupDate AS DateOfSale,
  InvoLine.CashSysID,
  InvoLine.InvoHeadID
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
    And (
      (InvoLine.Code) Between Forms!PrintInventoryReport!cboItemIDBegin
      And Forms!PrintInventoryReport!cboItemIDEnd
    )
    And (
      (InvoLine.TransType)= "S"
    )
  )
GROUP BY
  InvoLine.SetupDate,
  InvoLine.CashSysID,
  InvoLine.InvoHeadID
ORDER BY
  InvoLine.SetupDate;
