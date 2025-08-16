UPDATE
  PricingWorkTable
SET
  PricingWorkTable.NewPrice = [newprice] + 0.01
WHERE
  (
    (
      (
        Right(
          Format([newprice], "Fixed"),
          1
        )
      )= 4
    )
  );
