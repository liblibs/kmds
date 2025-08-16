SELECT
  ClubOrderXL.line AS LineID,
  ClubOrderXL.ItemID,
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
  ClubOrderXL.ClubID,
  ClubOrderXL.ClubName,
  ClubOrderXL.StoreName
FROM
  ClubOrderXL
  INNER JOIN Item ON ClubOrderXL.ItemID = Item.ID;
