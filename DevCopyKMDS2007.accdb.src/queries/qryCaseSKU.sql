SELECT
  ItemCorv1.sku AS Expr1,
  ItemCorv1.[case q] AS Expr2,
  ItemCorv1.appendSKU AS Expr3,
  [SKU] - 1 AS PackSKU
FROM
  ItemCorv1
WHERE
  (
    (
      ([ItemCorv1].[appendSKU])<> "N"
    )
  );
