INSERT INTO upcitem (ID, ItemID)
SELECT
  checkdigit11(
    Right([mfg part], 11)
  ) AS UPCCode,
  ItemCorv1.sku
FROM
  ItemCorv1
  INNER JOIN Item ON ItemCorv1.sku = Item.ID
WHERE
  (
    (
      (
        Len([mfg part])
      )= 12
    )
    AND (
      (ItemCorv1.[vendor number])= 100
    )
  );
