SELECT
  qryMetrologicScanBatchTotalByItemID.ItemID,
  qryMetrologicScanBatchTotalByItemID.SumOfQuantity,
  vwItemsWithCurrentPricing.Cost,
  vwItemsWithCurrentPricing.Price,
  vwItemsWithCurrentPricing.UnitsPerCase
FROM
  qryMetrologicScanBatchTotalByItemID
  INNER JOIN vwItemsWithCurrentPricing ON qryMetrologicScanBatchTotalByItemID.ItemID = vwItemsWithCurrentPricing.ID
ORDER BY
  qryMetrologicScanBatchTotalByItemID.ItemID;
