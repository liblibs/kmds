UPDATE
  Item
SET
  Item.Cost = [ITEM].[PRICE]
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
    )
  );
