UPDATE
  PSCInput
  INNER JOIN vwUPCItemsWithCurrentPricing ON PSCInput.Code = vwUPCItemsWithCurrentPricing.UPCCode
SET
  PSCInput.ItemID = vwUPCItemsWithCurrentPricing.id,
  PSCInput.Description = vwUPCItemsWithCurrentPricing.description,
  PSCInput.OnShelf = vwUPCItemsWithCurrentPricing.onshelf,
  PSCInput.Price = vwUPCItemsWithCurrentPricing.curprice,
  PSCInput.Cost = vwUPCItemsWithCurrentPricing.cost,
  PSCInput.SalesDepartmentID = vwUPCItemsWithCurrentPricing.SalesDepartmentID,
  PSCInput.Locked = True;
