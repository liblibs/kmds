SELECT
  Max(UPCItem.ID) AS MaxOfID,
  UPCItem.ItemID
FROM
  UPCItem
WHERE
  (
    (
      (
        Len([id])
      )= 12
    )
  )
GROUP BY
  UPCItem.ItemID;
