INSERT INTO UPCItem (ID, ItemID)
SELECT
  upcatoe([id]) AS UPCE,
  UPCItem.ItemID
FROM
  UPCItem
WHERE
  (
    (
      (
        Len([id])
      )= 10
    )
    AND (
      (
        Len(
          upcatoe([id])
        )
      )= 8
    )
  );
