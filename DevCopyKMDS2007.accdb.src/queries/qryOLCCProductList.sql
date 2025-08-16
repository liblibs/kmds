SELECT
  ProductListFromOLCCWebsite.F1 AS Expr1,
  ProductListFromOLCCWebsite.F3 AS Expr2,
  CLng(
    Left(
      [f1],
      Len([f1])-1
    )& l2n(
      Right([f1], 1)
    )
  ) AS Liq,
  Right([f1], 1) AS [Size] INTO OLCCProd
FROM
  ProductListFromOLCCWebsite;
