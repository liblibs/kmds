SELECT
  MeadowlandInventoryReport13.NewUPC1 AS Expr1,
  Len([NewUPC1]) AS lid,
  Left([NewUPC1], 1) AS le,
  MeadowlandInventoryReport13.NewItemID AS Expr2
FROM
  MeadowlandInventoryReport13
WHERE
  (
    (
      (
        Len([NewUPC1])
      )= 12
    )
    And (
      (
        Left([NewUPC1], 1)
      )= 0
    )
    And (
      (
        MeadowlandInventoryReport13.NewItemID
      )> 0
    )
  );
