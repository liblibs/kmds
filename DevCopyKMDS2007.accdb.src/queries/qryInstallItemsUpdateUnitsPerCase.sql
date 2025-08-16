UPDATE
  Item
  INNER JOIN [Size] ON Item.SizeID = Size.ID
SET
  Item.UnitsPerCase = [CaseQty]
WHERE
  (
    (
      (Item.UnitsPerCase)= 0
    )
    AND (
      (Item.SalesDepartmentID)= 1
    )
  );
