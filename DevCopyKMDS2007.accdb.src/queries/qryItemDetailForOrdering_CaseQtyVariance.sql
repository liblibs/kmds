SELECT
  Item.ID,
  Item.Description,
  Item.UnitsPerCase AS MyCaseQty,
  OLASItemDetailForOrdering.BottlesPerCase AS CaseQtyOLAS
FROM
  Item
  INNER JOIN OLASItemDetailForOrdering ON Item.ID = OLASItemDetailForOrdering.NewItemCode
WHERE
  (
    (
      (Item.UnitsPerCase)<> [OLASItemDetailForOrdering].[BottlesPerCase]
    )
  );
