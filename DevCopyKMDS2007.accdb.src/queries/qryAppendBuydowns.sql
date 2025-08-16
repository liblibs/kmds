PARAMETERS parBuyDownID Short; INSERT INTO BuyDownItem (
  BuyDownID, Line, [Order], ItemsCode,
  QuantitySold, DiscountDollars, SpcPrice
)
SELECT
  [parBuyDownID] AS BuyDownID,
  ConvCIGSLINE.Line AS Expr1,
  ConvCIGSLINE.Order AS Expr2,
  ConvCIGSLINE.Item AS ItemsCode,
  ConvCIGSLINE.ShpQty AS Expr3,
  ConvCIGSLINE.ShpDollars AS Expr4,
  ConvCIGSLINE.SpcPrice AS Expr5
FROM
  ConvCIGSLINE,
  Item
  LEFT JOIN BuyDownItem ON Item.ID = BuyDownItem.ItemsCode
WHERE
  (
    (
      ([ConvCIGSLINE].[Order])= [parOrder]
    )
    AND (
      (BuyDownItem.ItemsCode) Is Null
    )
  )
ORDER BY
  ConvCIGSLINE.Line;
