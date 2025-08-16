UPDATE
  SalesDepartment,
  InvoLine5
  INNER JOIN Item ON InvoLine5.Code = Item.ID
SET
  InvoLine5.Cost = (involine5.price)*(1 - [TargetGP])
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
    )
    AND (
      (InvoLine5.TransType)= "C"
    )
    AND (
      (SalesDepartment.ID)= 2
    )
  );
