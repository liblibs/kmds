UPDATE
  PricingWorkTable
SET
  PricingWorkTable.NewPrice = [newprice] + 0.03
WHERE
  (
    (
      (
        Right(
          Format([newprice], "Fixed"),
          1
        )
      )= 2
    )
  );
