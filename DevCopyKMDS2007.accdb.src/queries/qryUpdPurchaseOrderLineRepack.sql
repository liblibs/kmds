UPDATE
  Item
  INNER JOIN PurchaseOrderLine ON Item.ID = PurchaseOrderLine.ItemID
SET
  PurchaseOrderLine.Repack = IIf(
    [QuantityOrdered] -(
      (
        [quantityOrdered] \[UnitsPerCase]
      )* [UnitsPerCase]
    )= 0,
    0,
    -1
  );
