SELECT
  UPCItem.ItemID,
  Sum(1) AS [Count]
FROM
  UPCItem
WHERE
  (
    (
      (UPCItem.ID)<> CStr([ItemID])
    )
  )
GROUP BY
  UPCItem.ItemID
HAVING
  (
    (
      (
        Sum(1)
      )> 1
    )
  )
ORDER BY
  UPCItem.ItemID;
