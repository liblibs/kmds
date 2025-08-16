SELECT
  MeadowlandInventoryReport13.NewUPC2 AS Expr1,
  Len([NewUPC2]) AS lid,
  Left([NewUPC2], 3) AS le,
  MeadowlandInventoryReport13.NewItemID AS Expr2
FROM
  MeadowlandInventoryReport13
WHERE
  (
    (
      (
        Len([NewUPC2])
      )= 11
    )
    AND (
      (
        Left([NewUPC2], 3)
      )<> "000"
    )
  );
