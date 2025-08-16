SELECT
  SalesDepartment.ID,
  SalesDepartment.Description,
  Sum(InvoLine.Quantity) AS SumOfQuantity,
  Sum(InvoLine.Price) AS SumOfPrice
FROM
  SalesDepartment
  RIGHT JOIN (
    InvoLine
    LEFT JOIN Item ON InvoLine.Code = Item.ID
  ) ON SalesDepartment.ID = Item.SalesDepartmentID
WHERE
  (
    (
      (InvoLine.TransType)<> 'T'
      And (InvoLine.TransType)<> 'X'
    )
  )
GROUP BY
  SalesDepartment.ID,
  SalesDepartment.Description;
