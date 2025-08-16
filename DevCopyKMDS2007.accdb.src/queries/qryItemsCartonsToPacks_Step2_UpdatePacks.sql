UPDATE
  (
    Item
    INNER JOIN UPCItem ON Item.SubPack = UPCItem.ID
  )
  INNER JOIN Item AS Item_1 ON UPCItem.ItemID = Item_1.ID
SET
  Item_1.Onhand = Item_1.Onhand +(Item.onhand * Item.UnitsPerPack)
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
  );
