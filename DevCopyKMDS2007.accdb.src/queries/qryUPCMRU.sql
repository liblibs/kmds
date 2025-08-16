SELECT
  First(UPCItem.ChangeDate) AS FirstOfChangeDate,
  UPCItem.ItemID,
  UPCItem.ID,
  Len([id]) AS lid
FROM
  UPCItem
GROUP BY
  UPCItem.ItemID,
  UPCItem.ID,
  Len([id])
HAVING
  (
    (
      (
        Len([id])
      )= 8
      Or (
        Len([id])
      )= 12
    )
  )
ORDER BY
  First(UPCItem.ChangeDate) DESC;
