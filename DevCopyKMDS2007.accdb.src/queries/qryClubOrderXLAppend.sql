PARAMETERS parOrderNumber Long; INSERT INTO OrderLine (
  OrdheadID, LineID, ItemID, QuantityOrdered,
  QuantityShipped, Price, Cost
)
SELECT
  [parOrderNumber] AS OrdheadID,
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
  Item.Cost
FROM
  ClubOrderXL
  INNER JOIN Item ON ClubOrderXL.ItemID = Item.ID
WHERE
  (
    (
      (
        IIf(
          IsNumeric([Qty]),
          [Qty],
          0
        )
      ) Is Not Null
      And (
        IIf(
          IsNumeric([Qty]),
          [Qty],
          0
        )
      ) Is Not Null
    )
  );
