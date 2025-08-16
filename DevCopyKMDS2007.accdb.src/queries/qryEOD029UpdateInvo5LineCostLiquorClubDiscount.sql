UPDATE
  InvoLine5
  INNER JOIN Item ON InvoLine5.Code = Item.ID
SET
  InvoLine5.Cost = InvoLine5.Price
WHERE
  (
    (
      (Item.SalesDepartmentID)= 2
    )
  );
