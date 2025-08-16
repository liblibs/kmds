INSERT INTO ItemSalePrice (
  ItemID, SalePrice, StartDate, EndDate,
  OldPrice
)
SELECT
  ConvITEMS.Code AS Expr1,
  ConvITEMS.NewPrice AS Expr2,
  ([Starts]) AS st,
  ConvITEMS.Ends AS Expr3,
  ConvITEMS.Price AS Expr4
FROM
  ConvITEMS
WHERE
  (
    (
      (
        ([Starts])
      )> #5/31/2004#
    )
    And (
      (ConvITEMS.Ends)> "#05/31/2004#"
    )
  )
  Or (
    (
      (
        ([Starts])
      ) Is Not Null
    )
    And (
      (ConvITEMS.Ends) Is Not Null
    )
  );
