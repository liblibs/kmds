SELECT
  ClubOrderCSV.line AS LineID,
  ClubOrderCSV.ItemID,
  IIf(
    IsNumeric([Qty]),
    [Qty],
    0
  ) AS QuantityOrdered,
  IIf(
    IsNumeric([Qty]),
    [Qty],
    0
  ) AS QuantityShipped,
  Item.Price,
  Item.Cost,
  ClubOrderCSV.ClubID,
  ClubOrderCSV.ClubName,
  ClubOrderCSV.StoreName,
  ClubOrderCSV.QuantityOrdered,
  ClubOrderCSV.QuantityOrdered AS QuantityShipped
FROM
  ClubOrderCSV
  INNER JOIN Item ON ClubOrderCSV.ItemId = Item.ID;
