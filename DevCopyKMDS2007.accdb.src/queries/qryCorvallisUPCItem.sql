INSERT INTO upcitem (ID, ItemID)
SELECT
  checkdigit11(
    Right([scan number], 11)
  ) AS UPCCode,
  ItemCorv1.sku
FROM
  ItemCorv1
  INNER JOIN Item ON ItemCorv1.sku = Item.ID
WHERE
  (
    (
      (
        Len([scan number])
      )= 13
    )
  );
