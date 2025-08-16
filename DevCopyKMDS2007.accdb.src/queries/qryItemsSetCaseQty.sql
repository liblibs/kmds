UPDATE
  Item
  INNER JOIN ItemCaseQty ON Item.ID = ItemCaseQty.Code
SET
  Item.UnitsPerCase = [CaseQty]
WHERE
  (
    (
      (Item.UnitsPerCase)<> [CaseQty]
    )
  );
