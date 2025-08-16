SELECT
  OrderLine.OrdheadID,
  OrderLine.LineID,
  OrderLine.ID
FROM
  OrderLine
ORDER BY
  OrderLine.LineID DESC;
