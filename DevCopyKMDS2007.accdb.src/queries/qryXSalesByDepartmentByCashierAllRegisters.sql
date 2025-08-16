SELECT
  SalesDepartment.ID,
  SalesDepartment.Description,
  Sum(InvoLineAllRegisters.Quantity) AS SumOfQuantity,
  Sum(InvoLineAllRegisters.Price) AS SumOfPrice,
  Cashier.ID,
  Cashier.Name
FROM
  (
    (
      InvoLineAllRegisters
      INNER JOIN Item ON InvoLineAllRegisters.Code = Item.ID
    )
    INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
  )
  INNER JOIN Cashier ON InvoLineAllRegisters.CashierID = Cashier.ID
WHERE
  (
    (
      (InvoLineAllRegisters.TransType)<> 'T'
      And (InvoLineAllRegisters.TransType)<> 'X'
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
