UPDATE
  Item
SET
  Item.Location = "Office"
WHERE
  (
    (
      (Item.UnitsPerOrder)>= [Unitspercase]
    )
    AND (
      (Item.UnitsPerCase)> 1
    )
    AND (
      (Item.SalesDepartmentID)= 1
    )
    AND (
      (Item.SizeID)= 20
      Or (Item.SizeID)= 37
      Or (Item.SizeID)= 5
    )
  );
