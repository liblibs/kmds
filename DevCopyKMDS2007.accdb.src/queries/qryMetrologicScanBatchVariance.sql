SELECT
  vwItemsWithCurrentPricing.SalesDptDescr,
  vwItemsWithCurrentPricing.CategoryDescr,
  vwItemsWithCurrentPricing.ID,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.PkgSize,
  qryMetrologicScanBatchTotalByItemID.SumOfQuantity AS [Count],
  vwItemsWithCurrentPricing.OnShelf,
  [SumOfQuantity] - [OnShelf] AS Variance,
  [curPrice] *([SumOfQuantity] - [OnShelf]) AS [Variance Dollars],
  qryMetrologicScanBatchTotalByItemID.FirstOfCashierID AS CashierID,
  qryMetrologicScanBatchTotalByItemID.InvoLineID
FROM
  qryMetrologicScanBatchTotalByItemID
  INNER JOIN vwItemsWithCurrentPricing ON qryMetrologicScanBatchTotalByItemID.ItemID = vwItemsWithCurrentPricing.ID
ORDER BY
  vwItemsWithCurrentPricing.SalesDptDescr,
  vwItemsWithCurrentPricing.CategoryDescr,
  vwItemsWithCurrentPricing.ID;
