INSERT INTO Audit (
  ItemID, PhysicalCount, OnShelf, Price,
  Cost, VarianceAmount, CountFront
)
SELECT
  CrescInventory.ID,
  CrescInventory.SumOfQuantity,
  Item.Onhand,
  Item.Price,
  Item.Cost,
  [sumofquantity] - [onhand] AS va,
  CrescInventory.sumofquantity
FROM
  CrescInventory
  INNER JOIN Item ON CrescInventory.ID = Item.ID
WHERE
  (
    (
      ([sumofquantity] - [onhand])< 0
    )
  );
