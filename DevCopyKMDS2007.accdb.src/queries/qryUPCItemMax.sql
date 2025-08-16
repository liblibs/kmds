SELECT
  UPCItem.ItemID,
  Max(UPCItem.ID) AS UPCCodeMax,
  Sum(1) AS [Count]
FROM
  UPCItem
WHERE
  (
    (
      (
        Len([ID])
      )> 7
    )
    AND (
      (UPCItem.ID) Not Like "99900*"
    )
  )
GROUP BY
  UPCItem.ItemID
ORDER BY
  UPCItem.ItemID,
  Max(UPCItem.ID);
