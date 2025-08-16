PARAMETERS parOrdheadID Long;
SELECT
  Max(OrderLine.LineID) AS MaxLineID
FROM
  OrderLine
WHERE
  (
    (
      (OrderLine.OrdheadID)= [parOrdheadID]
    )
  );
