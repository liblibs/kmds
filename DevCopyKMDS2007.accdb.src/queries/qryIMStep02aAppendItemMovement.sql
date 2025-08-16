INSERT INTO ItemMovement (
  Code, Description, CurrMM, YTDQty,
  Jan, Feb, Mar, Apr, May, Jun, Jul, Aug,
  Sep, Oct, Nov, [Dec]
)
SELECT
  qryItemMovement.Code,
  TranTypes.Description,
  0 AS CurrMM,
  Sum(qryItemMovement.SumOfQuantity) AS YTDQty,
  Sum(
    IIf([MM] = "01", [SumOfQuantity], 0)
  ) AS Jan,
  Sum(
    IIf([MM] = "02", [SumOfQuantity], 0)
  ) AS Feb,
  Sum(
    IIf([MM] = "03", [SumOfQuantity], 0)
  ) AS Mar,
  Sum(
    IIf([MM] = "04", [SumOfQuantity], 0)
  ) AS Apr,
  Sum(
    IIf([MM] = "05", [SumOfQuantity], 0)
  ) AS May,
  Sum(
    IIf([MM] = "06", [SumOfQuantity], 0)
  ) AS Jun,
  Sum(
    IIf([MM] = "07", [SumOfQuantity], 0)
  ) AS Jul,
  Sum(
    IIf([MM] = "08", [SumOfQuantity], 0)
  ) AS Aug,
  Sum(
    IIf([MM] = "09", [SumOfQuantity], 0)
  ) AS Sep,
  Sum(
    IIf([MM] = "10", [SumOfQuantity], 0)
  ) AS Oct,
  Sum(
    IIf([MM] = "11", [SumOfQuantity], 0)
  ) AS Nov,
  Sum(
    IIf([MM] = "12", [SumOfQuantity], 0)
  ) AS [Dec]
FROM
  qryItemMovement
  LEFT JOIN TranTypes ON qryItemMovement.TransType = TranTypes.Code
GROUP BY
  qryItemMovement.Code,
  TranTypes.Description,
  0
ORDER BY
  TranTypes.Description;
