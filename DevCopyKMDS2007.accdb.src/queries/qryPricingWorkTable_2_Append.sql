INSERT INTO PricingWorkTable (NewPrice, NewCost, NewGP)
SELECT
  qryItemsByVendor.*,
  qryItemsByVendor.Price AS Expr1,
  qryItemsByVendor.Cost AS Expr2,
  IIf(
    ([price] > 0),
    (
      ([Price] - [cost])/ [Price]
    ),
    0
  ) AS NewGP
FROM
  qryItemsByVendor
WHERE
  (
    (
      (qryItemsByVendor.Price)> 0
    )
    And (
      (qryItemsByVendor.VendorID)= Forms!PricingToolSelect!ID
    )
  );
