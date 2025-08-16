UPDATE
  PricingWorkTable
SET
  PricingWorkTable.NewGP = (
    [PricingWorkTable].[NewPrice] - [PricingWorkTable].[Cost]
  )/ [PricingWorkTable].[NewPrice]
WHERE
  (
    (
      (PricingWorkTable.NewPrice)> 0
    )
  );
