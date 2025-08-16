UPDATE
  (
    Item
    INNER JOIN UPCItem ON Item.SubPack = UPCItem.ID
  )
  INNER JOIN Item AS Item_1 ON UPCItem.ItemID = Item_1.ID
SET
  Item.Onhand = 0
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
