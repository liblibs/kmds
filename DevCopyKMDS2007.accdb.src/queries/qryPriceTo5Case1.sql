UPDATE
  PricingWorkTable
SET
  PricingWorkTable.NewPrice = [newprice] + 0.04
WHERE
  (
    (
      (
        Right(
          Format([newprice], "Fixed"),
          1
        )
      )= 1
    )
  );
