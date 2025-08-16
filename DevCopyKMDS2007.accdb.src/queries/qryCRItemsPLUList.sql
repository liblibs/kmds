SELECT
  Item.ID,
  Item.Description,
  Item.PLUList,
  Item.Price
FROM
  Item
WHERE
  (
    (
      (Item.PLUList)<> 0
    )
  )
ORDER BY
  Item.Description;
