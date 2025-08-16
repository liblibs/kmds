SELECT
  Item.SalesDepartmentID,
  UPCItem.ID,
  upcatoe([upcitem].[id]) AS newUPC,
  UPCItem.ItemID,
  Len(
    upcatoe([upcitem].[id])
  ) AS lid INTO NewUPC
FROM
  UPCItem
  INNER JOIN Item ON UPCItem.ItemID = Item.ID
WHERE
  (
    (
      (Item.SalesDepartmentID)> 2
    )
  );
