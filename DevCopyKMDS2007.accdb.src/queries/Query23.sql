INSERT INTO UPCItem (ItemID, ID)
SELECT
  Item.ID AS Item,
  checkdigit11("0" & [UPcitem].[id]) AS NewID
FROM
  UPCItem
  INNER JOIN Item ON UPCItem.ItemID = Item.ID
WHERE
  (
    (
      (
        Len([upcitem].[id])
      )= 10
    )
  )
ORDER BY
  Item.ID;
