SELECT
  UPCItem.ItemID,
  Count(UPCItem.ItemID) AS CountOfItemID
FROM
  UPCItem
GROUP BY
  UPCItem.ItemID
HAVING
  (
    (
      (
        Count(UPCItem.ItemID)
      )= 1
    )
  );
