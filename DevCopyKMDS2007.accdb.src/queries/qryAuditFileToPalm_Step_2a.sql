SELECT
  UPCItem.ItemID,
  First(UPCItem.ChangeDate) AS FirstOfChangeDate,
  First(UPCItem.ID) AS FirstOfID
FROM
  UPCItem
WHERE
  (
    (
      (
        Len([ID])
      )> 7
    )
  )
GROUP BY
  UPCItem.ItemID
ORDER BY
  UPCItem.ItemID,
  First(UPCItem.ChangeDate) DESC;
