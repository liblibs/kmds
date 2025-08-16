PARAMETERS parPriceGroupID Short,
parPrice Currency,
parQtyPack Short,
parQtyPackPrice Currency,
parCost Currency;
UPDATE
  Item
SET
  Item.Price = [parPrice],
  Item.QtyPack = [parQtyPack],
  Item.QtyPackPrice = [parQtyPackPrice],
  Item.ChangeDate = Now(),
  Item.Cost = [parCost]
WHERE
  (
    (
      (Item.PriceGroupID)= [parPriceGroupID]
    )
    AND (
      (Item.SalesDepartmentID)<> 1
    )
  );
