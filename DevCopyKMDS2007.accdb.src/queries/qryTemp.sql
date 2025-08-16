SELECT
  Items.Code AS Expr1,
  Items.Description AS Expr2,
  Items.Onhand AS Expr3,
  Items.Price AS Expr4,
  Items.PriceGroupID AS Expr5,
  Items.SalesDepartmentID AS Expr6,
  Items.StartDate AS Expr7,
  Items.EndDate AS Expr8,
  Items.NewPrice AS Expr9
FROM
  Items
WHERE
  (
    (
      ([Items].[Description]) Like 'M*'
    )
  )
ORDER BY
  Items.Description;
