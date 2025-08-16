PARAMETERS parOrdheadID Long,
parItemID Decimal;
SELECT
  OrderLine.OrdheadID,
  OrderLine.LineID,
  OrderLine.ItemID
FROM
  OrderLine
WHERE
  (
    (
      (OrderLine.OrdheadID)= [parOrdheadID]
    )
    AND (
      (OrderLine.ItemID)= [parItemID]
    )
  );
