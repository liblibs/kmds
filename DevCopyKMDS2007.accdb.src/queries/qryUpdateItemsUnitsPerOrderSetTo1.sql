UPDATE
  Item
  INNER JOIN [Size] ON Item.SizeID = Size.ID
SET
  Item.UnitsPerOrder = 1
WHERE
  (
    (
      (Item.UnitsPerOrder)= 0
    )
    AND (
      (Size.CaseQty)> [OrderTo]
    )
  );
