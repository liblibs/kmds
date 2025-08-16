INSERT INTO OrderLine (
  OrdheadID, LineID, ItemId, QuantityOrdered,
  QuantityShipped, Price, Cost
)
SELECT
  [parOrderNumber] AS OrdheadID,
  ClubOrderCSV.LineCount AS LineID,
  ClubOrderCSV.ItemId,
  ClubOrderCSV.QuantityOrdered,
  ClubOrderCSV.QuantityOrdered,
  Item.Price,
  Item.Cost
FROM
  ClubOrderCSV
  INNER JOIN Item ON ClubOrderCSV.ItemId = Item.ID
WHERE
  (
    (
      (ClubOrderCSV.LineCount)> 0
    )
    AND (
      (ClubOrderCSV.QuantityOrdered)<> 0
    )
  );
