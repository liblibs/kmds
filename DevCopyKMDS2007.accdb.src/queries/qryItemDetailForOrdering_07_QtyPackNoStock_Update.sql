UPDATE
  Item
  INNER JOIN OLASItemDetailForOrdering ON Item.ID = OLASItemDetailForOrdering.NewItemCode
SET
  Item.QtyPack = [OLASItemDetailForOrdering].[QuantityAvailable];
