UPDATE
  PricingWorkTable
SET
  PricingWorkTable.NewPrice = [newprice] + 0.02
WHERE
  (
    (
      (
        Right(
          Format([newprice], "Fixed"),
          1
        )
      )= 7
    )
  );
