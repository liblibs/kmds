PARAMETERS parItemID Long,
parQuantity Short,
parDollars Currency;
UPDATE
  BuyDownItem
SET
  BuyDownItem.QuantitySold = BuyDownItem.QuantitySold + [parQuantity],
  BuyDownItem.DiscountDollars = BuyDownItem.DiscountDollars + [parDollars]
WHERE
  (
    (
      (BuyDownItem.ItemsCode)= [parItemID]
    )
  );
