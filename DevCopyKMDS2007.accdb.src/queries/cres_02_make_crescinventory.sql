SELECT
  Item.ID,
  Item.SalesDepartmentID,
  Sum(InvoLine.Quantity) AS SumOfQuantity INTO CrescInventory
FROM
  InvoLine
  INNER JOIN Item ON InvoLine.Code = Item.ID
WHERE
  (
    (
      (InvoLine.SetupDate)> #3/9/2017#
    )
    AND (
      (InvoLine.TransType)= "S"
    )
  )
GROUP BY
  Item.ID,
  Item.SalesDepartmentID
HAVING
  (
    (
      (Item.SalesDepartmentID)> 21
    )
  )
ORDER BY
  Item.ID;
