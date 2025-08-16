UPDATE
  CRES_ITEM
  INNER JOIN Item ON CRES_ITEM.ID = Item.ID
SET
  Item.Onhand = [CRES_ITEM].[Onhand]
WHERE
  (
    (
      (Item.SalesDepartmentID)> 21
    )
  );
