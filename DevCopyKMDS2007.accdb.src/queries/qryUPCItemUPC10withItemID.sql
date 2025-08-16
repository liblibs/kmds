SELECT
  UPCItem.ID,
  UPCItem.ItemID,
  Len([upcitem].[id]) AS lid,
  Item.Description,
  Item.SalesDepartmentID,
  Mid([UPCITEM].[ID], 2, 10) AS upc10
FROM
  UPCItem
  INNER JOIN Item ON UPCItem.ItemID = Item.ID
WHERE
  (
    (
      (
        Len(upcitem.id)
      )= 12
    )
    And (
      (Item.SalesDepartmentID)= 1
    )
  )
ORDER BY
  Len([upcitem].[id]);
