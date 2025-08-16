UPDATE
  Item
  INNER JOIN OLASItemDetailForOrdering ON Item.ID = OLASItemDetailForOrdering.NewItemCode
SET
  Item.UnitsPerOrder = [OLASItemDetailForOrdering].[MinOrderQty]
WHERE
  (
    (
      (Item.UnitsPerOrder)<> OLASItemDetailForOrdering.MinOrderQty
    )
    And (
      (
        OLASItemDetailForOrdering.MinOrderQty
      )> 0
    )
  );
