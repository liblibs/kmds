UPDATE
  InvoLine5
  INNER JOIN Item ON InvoLine5.Code = Item.ID
SET
  InvoLine5.Cost = Item.Cost * InvoLine5.quantity
WHERE
  (
    (
      (Item.SalesDepartmentID)> 2
    )
  );
