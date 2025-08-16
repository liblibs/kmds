SELECT
  OrdHead.ID,
  OrdHead.OrdDate,
  Club.ID,
  Club.Name,
  OrderLine.ItemID,
  OrderLine.QuantityShipped,
  OrderLine.Price,
  OrderLine.Cost,
  OrderLine.SetupDate,
  OrderLine.ChangeDate,
  [QuantityShipped] * [Price] AS ExtPrice,
  [QuantityShipped] * [Cost] AS ExtCost
FROM
  (
    OrderLine
    INNER JOIN OrdHead ON OrderLine.OrdheadID = OrdHead.ID
  )
  INNER JOIN Club ON OrdHead.ClubID = Club.ID
WHERE
  (
    (
      (OrderLine.QuantityShipped)<> 0
    )
  );
