UPDATE
  (
    InvoLine5
    INNER JOIN Item ON InvoLine5.Code = Item.ID
  )
  INNER JOIN SalesDepartment ON Item.SalesDepartmentID = SalesDepartment.ID
SET
  InvoLine5.Cost = (involine5.price)*(1 - [TargetGP])
WHERE
  (
    (
      (InvoLine5.Cost)= 0
    )
    AND (
      (Item.SalesDepartmentID)> 2
    )
  );
