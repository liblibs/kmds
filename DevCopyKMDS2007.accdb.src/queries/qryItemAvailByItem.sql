PARAMETERS parItemID Long;
SELECT
  vwItemsWithCurrentPricing.ID,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.OnShelf,
  vwItemsWithCurrentPricing.CurPrice,
  vwItemsWithCurrentPricing.BuyDown,
  vwItemsWithCurrentPricing.SalesDepartmentID,
  vwItemsWithCurrentPricing.Price,
  vwItemsWithCurrentPricing.Reserved
FROM
  vwItemsWithCurrentPricing
WHERE
  (
    (
      (vwItemsWithCurrentPricing.ID)= [parItemID]
    )
  );
