UPDATE
  Item
  INNER JOIN OLASItemDetailForOrdering ON Item.ID = OLASItemDetailForOrdering.NewItemCode
SET
  Item.UnitsPerCase = [OLASItemDetailForOrdering].[BottlesPerCase]
WHERE
  (
    (
      (Item.UnitsPerCase)<> [OLASItemDetailForOrdering].[BottlesPerCase]
    )
  );
