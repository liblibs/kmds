SELECT
  qryKitsOnFile.ID,
  qryKitsOnFile.ItemID AS KitItemID,
  qryKitsOnFile.Description AS KitDescription,
  qryKitsOnFile.Price AS KitPrice,
  qryKitsOnFile.RetailPrice,
  IIf(
    [RetailPrice] = 0, 0, [Discount] / [RetailPrice]
  ) AS KitDiscPerct,
  qryKitsOnFile.Discount AS KitDiscount,
  KitItem.ItemID AS PartID,
  Item_1.Description AS PartDescription,
  KitItem.Price AS PartPrice,
  KitItem.SetupDate,
  KitItem.ChangeDate
FROM
  (
    KitItem
    RIGHT JOIN qryKitsOnFile ON KitItem.KitID = qryKitsOnFile.ID
  )
  LEFT JOIN Item AS Item_1 ON KitItem.ItemID = Item_1.ID;
