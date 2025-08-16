SELECT
  Format(
    [InvoLineAllRegisters]![SetupDate],
    "mm-dd    hh:""00 ""ampm"
  ) AS DateHour,
  Sum(InvoLineAllRegisters.Quantity) AS SumOfQuantity,
  Sum(InvoLineAllRegisters.Price) AS SumOfPrice
FROM
  InvoLineAllRegisters
  INNER JOIN (
    Item
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  ) ON InvoLineAllRegisters.Code = Item.ID
WHERE
  (
    (
      (InvoLineAllRegisters.TransType)<> 'T'
      And (InvoLineAllRegisters.TransType)<> 'X'
    )
    AND (
      (SalesDepartment.ID)<= 50
    )
  )
GROUP BY
  Format(
    [InvoLineAllRegisters]![SetupDate],
    "mm-dd    hh:""00 ""ampm"
  )
ORDER BY
  Format(
    [InvoLineAllRegisters]![SetupDate],
    "mm-dd    hh:""00 ""ampm"
  );
