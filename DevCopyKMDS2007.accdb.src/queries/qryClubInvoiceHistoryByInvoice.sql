SELECT
  InvoLine.TransType,
  InvoLine.InvoHeadID,
  InvoLine.CashSysID,
  InvoLine.SetupDate,
  InvoLine.SalesDept AS ClubID,
  Sum(InvoLine.Price) AS SumOfPrice
FROM
  InvoLine
  INNER JOIN Club ON InvoLine.SalesDept = Club.ID
WHERE
  (
    (
      (Club.Name) Between [Forms]![PrintInventoryReport]![cboClubIDBegin]
      And [Forms]![PrintInventoryReport]![cboClubIDEnd]
    )
  )
GROUP BY
  InvoLine.TransType,
  InvoLine.InvoHeadID,
  InvoLine.CashSysID,
  InvoLine.SetupDate,
  InvoLine.SalesDept
HAVING
  (
    (
      (InvoLine.TransType)= "C"
    )
    And (
      (InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
  );
