UPDATE
  Item
  INNER JOIN [Size] ON Item.SizeID = Size.ID
SET
  Item.UnitsPerOrder = [CaseQty]
WHERE
  (
    (
      (Item.UnitsPerOrder)= 0
    )
    AND (
      (Size.CaseQty)<= [OrderTo]
    )
  );
