SELECT
  ClubOrderCSV.LineCount AS LineID,
  ClubOrderCSV.ItemID,
  Item.Price,
  Item.Cost,
  ClubOrderCSV.ClubID,
  Club.Name AS ClubName,
  ClubOrderCSV.QuantityOrdered,
  ClubOrderCSV.QuantityOrdered AS QuantityShipped
FROM
  (
    ClubOrderCSV
    INNER JOIN Item ON ClubOrderCSV.ItemId = Item.ID
  )
  INNER JOIN Club ON ClubOrderCSV.ClubID = Club.ID
WHERE
  (
    (
      (ClubOrderCSV.LineCount)> 0
    )
  );
