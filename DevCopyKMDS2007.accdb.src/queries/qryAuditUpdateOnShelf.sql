UPDATE
  Audit
  INNER JOIN vwItemsWithCurrentPricing ON Audit.ItemID = vwItemsWithCurrentPricing.ID
SET
  Audit.OnShelf = vwItemsWithCurrentPricing.onshelf,
  Audit.VarianceAmount = audit.physicalcount - vwItemsWithCurrentPricing.onshelf
WHERE
  (
    (
      (Audit.OnShelf)<> [vwItemsWithCurrentPricing].[onshelf]
    )
  );
