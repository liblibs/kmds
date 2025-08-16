SELECT
  ConvITEMS1.Code AS Expr1,
  ConvITEMS1.Price AS Expr2,
  ConvITEMS1.NewPrice AS Expr3,
  ConvITEMS1.Starts AS Expr4,
  ConvITEMS1.Ends AS Expr5,
  IIf(
    Not IsNull([Starts])
    And CDate([Starts])> #2/28/2005#,
    [NewPrice],
    [Price]
  ) AS csPrice
FROM
  ConvITEMS1
WHERE
  (
    (
      ([ConvITEMS1].[SalesDept])= 1
    )
  )
ORDER BY
  ConvITEMS1.Starts DESC;
