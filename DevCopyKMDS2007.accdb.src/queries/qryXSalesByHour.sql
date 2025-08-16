SELECT
  Format(
    [InvoLine]![SetupDate], "mm-dd    hh:""00 ""ampm"
  ) AS DateHour,
  Sum(InvoLine.Quantity) AS SumOfQuantity,
  Sum(InvoLine.Price) AS SumOfPrice
FROM
  (
    InvoLine
    INNER JOIN Item ON InvoLine.Code = Item.ID
  )
  INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
WHERE
  (
    (
      (InvoLine.TransType)<> 'T'
      And (InvoLine.TransType)<> 'X'
    )
    AND (
      (SalesDepartment.ID)<= 50
    )
  )
GROUP BY
  Format(
    [InvoLine]![SetupDate], "mm-dd    hh:""00 ""ampm"
  )
ORDER BY
  Format(
    [InvoLine]![SetupDate], "mm-dd    hh:""00 ""ampm"
  );
