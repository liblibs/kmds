UPDATE
  Item
SET
  Item.Location = "Back"
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
      (Item.SizeID)= 10
      Or (Item.SizeID)= 75
      Or (Item.SizeID)= 20
      Or (Item.SizeID)= 50
      Or (Item.SizeID)= 17
    )
  );
