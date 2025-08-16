SELECT
  SalesDepartment.ID,
  SalesDepartment.Description,
  Sum(InvoLine.Quantity) AS SumOfQuantity,
  Sum(InvoLine.Price) AS SumOfPrice,
  Cashier.ID,
  Cashier.Name
FROM
  Cashier
  INNER JOIN (
    (
      InvoLine
      INNER JOIN Item ON InvoLine.Code = Item.ID
    )
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  ) ON Cashier.ID = InvoLine.CashierID
WHERE
  (
    (
      (InvoLine.TransType)<> 'T'
      And (InvoLine.TransType)<> 'X'
    )
  )
GROUP BY
  SalesDepartment.ID,
  SalesDepartment.Description,
  Cashier.ID,
  Cashier.Name
HAVING
  (
    (
      (Cashier.ID)= [parCashierNumber]
    )
  )
ORDER BY
  SalesDepartment.ID,
  Cashier.ID;
