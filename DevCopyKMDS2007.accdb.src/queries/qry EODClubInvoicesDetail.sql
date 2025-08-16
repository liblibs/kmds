SELECT
  OrdHead.InvoiceNo,
  Club.ID,
  Club.Name,
  Item.ID,
  Item.Description,
  OrderLine.QuantityShipped,
  (
    [QuantityShipped] * Orderline.Price
  ) AS ExtPrice,
  OrdHead.DisDollars,
  OrdHead.InvDollars,
  [InvDollars] - [DisDollars] AS NetDollars,
  OrdHead.PaidFlag
FROM
  (
    OrderLine
    INNER JOIN (
      Club
      INNER JOIN OrdHead ON Club.ID = OrdHead.ClubID
    ) ON OrderLine.OrdheadID = OrdHead.ID
  )
  INNER JOIN Item ON OrderLine.ItemID = Item.ID
WHERE
  (
    (
      (OrderLine.QuantityShipped)<> 0
    )
    AND (
      (OrdHead.PaidFlag)<> 0
    )
  );
