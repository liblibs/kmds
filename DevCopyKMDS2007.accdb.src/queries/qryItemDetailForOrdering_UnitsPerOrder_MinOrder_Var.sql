SELECT
  Item.ID,
  Item.Description,
  Item.UnitsPerOrder,
  OLASItemDetailForOrdering.MinOrderQty
FROM
  Item
  INNER JOIN OLASItemDetailForOrdering ON Item.ID = OLASItemDetailForOrdering.NewItemCode
WHERE
  (
    (
      (Item.UnitsPerOrder)<> [OLASItemDetailForOrdering].[MinOrderQty]
    )
  );
