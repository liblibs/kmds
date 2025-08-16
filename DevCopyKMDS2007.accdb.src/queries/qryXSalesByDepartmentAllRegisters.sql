SELECT
  Item.SalesDepartmentID,
  SalesDepartment.Description,
  Sum(InvoLineAllRegisters.Quantity) AS SumOfQuantity,
  Sum(InvoLineAllRegisters.Price) AS SumOfPrice
FROM
  InvoLineAllRegisters
  INNER JOIN (
    SalesDepartment
    RIGHT JOIN Item ON SalesDepartment.ID = Item.SalesDepartmentID
  ) ON InvoLineAllRegisters.Code = Item.ID
WHERE
  (
    (
      (InvoLineAllRegisters.TransType)<> 'T'
      And (InvoLineAllRegisters.TransType)<> 'X'
    )
  )
GROUP BY
  Item.SalesDepartmentID,
  SalesDepartment.Description;
