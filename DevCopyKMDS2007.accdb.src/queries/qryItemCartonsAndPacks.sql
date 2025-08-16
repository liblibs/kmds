SELECT
  Item.ID AS CartonID,
  UPCItem.ID AS CartonUPC,
  Item.Description AS CartonDesc,
  Item.Price AS CartonPrice,
  Item.UnitsPerPack AS PacksPerCarton,
  Item_1.ID AS PackID,
  Item.SubPack AS PackUPC,
  Item_1.Description AS PackDesc,
  Item_1.Price AS PackPrice
FROM
  (
    Item
    INNER JOIN UPCItem ON Item.ID = UPCItem.ItemID
  )
  INNER JOIN (
    UPCItem AS UPCItem_1
    INNER JOIN Item AS Item_1 ON UPCItem_1.ItemID = Item_1.ID
  ) ON Item.SubPack = UPCItem_1.ID;
