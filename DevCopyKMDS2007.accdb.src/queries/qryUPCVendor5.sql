SELECT
  UPCItem.ID,
  UPCItem.ItemID,
  Len([id]) AS lid,
  Mid([id], 2, 5) AS Vendor5U
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
ORDER BY
  UPCItem.ID DESC;
