SELECT
  Item.ID,
  Item.Price,
  [01 Inventory Value By Code].Cost,
  [01 Inventory Value By Code].CurPrice,
  [01 Inventory Value By Code].ID,
  [01 Inventory Value By Code].Description
FROM
  [01 Inventory Value By Code]
  LEFT JOIN Item ON [01 Inventory Value By Code].ID = Item.ID
WHERE
  (
    (
      (Item.Price) Is Not Null
    )
  );
