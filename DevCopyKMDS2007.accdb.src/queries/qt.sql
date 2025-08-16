INSERT INTO ItemMovement (
  Code, Description, CurrMM, YTDQty,
  Jan, Feb, Mar, Apr, May, Jun, Jul, Aug,
  Sep, Oct, Nov, [Dec]
)
SELECT
  qryItemMovementByView.Code,
  TranTypes.Description,
  0 AS CurrMM,
  Sum(
    qryItemMovementByView.SumOfQuantity
  ) AS YTDQty,
  Sum(
    IIf([MM] = 1, [SumOfQuantity], 0)
  ) AS Jan,
  Sum(
    IIf([MM] = 2, [SumOfQuantity], 0)
  ) AS Feb,
  Sum(
    IIf([MM] = 3, [SumOfQuantity], 0)
  ) AS Mar,
  Sum(
    IIf([MM] = 4, [SumOfQuantity], 0)
  ) AS Apr,
  Sum(
    IIf([MM] = 5, [SumOfQuantity], 0)
  ) AS May,
  Sum(
    IIf([MM] = 6, [SumOfQuantity], 0)
  ) AS Jun,
  Sum(
    IIf([MM] = 7, [SumOfQuantity], 0)
  ) AS Jul,
  Sum(
    IIf([MM] = 8, [SumOfQuantity], 0)
  ) AS Aug,
  Sum(
    IIf([MM] = 9, [SumOfQuantity], 0)
  ) AS Sep,
  Sum(
    IIf([MM] = 10, [SumOfQuantity], 0)
  ) AS Oct,
  Sum(
    IIf([MM] = 11, [SumOfQuantity], 0)
  ) AS Nov,
  Sum(
    IIf([MM] = 12, [SumOfQuantity], 0)
  ) AS [Dec]
FROM
  qryItemMovementByView
  LEFT JOIN TranTypes ON qryItemMovementByView.TransType = TranTypes.Code
GROUP BY
  qryItemMovementByView.Code,
  TranTypes.Description,
  0
ORDER BY
  TranTypes.Description;
