SELECT
  InvoLine.Code,
  Sum(InvoLine.Quantity) AS SumOfQuantity,
  Sum(InvoLine.Price) AS SumOfPrice,
  Sum(InvoLine.Cost) AS SumOfCost
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.TransType)= "S"
      Or (InvoLine.TransType)= "C"
    )
    And (
      (InvoLine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
  )
GROUP BY
  InvoLine.Code;
