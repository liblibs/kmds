SELECT
  Item.ID AS CartonID,
  Item.Description AS CartonDesc,
  Item.Onhand AS CartonOnhand,
  (Item.UnitsPerPack * Item.Onhand) AS PacksSold,
  Item_1.ID AS PackID,
  Item_1.Description AS PackDesc,
  Item_1.Onhand,
  Item_1.Onhand +(Item.onhand * Item.UnitsPerPack) AS PackNewOnhand
FROM
  (
    Item
    INNER JOIN UPCItem ON Item.SubPack = UPCItem.ID
  )
  INNER JOIN Item AS Item_1 ON UPCItem.ItemID = Item_1.ID
WHERE
  (
    (
      (Item.Onhand)<> 0
    )
    And (
      (
        Len(upcitem.ID)
      )> 3
    )
  )
ORDER BY
  Item.Description;
