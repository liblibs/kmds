SELECT
  OrderLine.ItemID,
  Sum(OrderLine.QuantityShipped) AS OnReserve
FROM
  OrderLine
GROUP BY
  OrderLine.ItemID;
