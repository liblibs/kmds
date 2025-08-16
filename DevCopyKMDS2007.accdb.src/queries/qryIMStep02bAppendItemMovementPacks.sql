INSERT INTO ItemMovement (
  Description, Code, YTDQty, CurrMM,
  Jan, Feb, Mar, Apr, May, Jun, Jul, Aug,
  Sep, Oct, Nov, [Dec]
)
SELECT
  " Sales - Packs" AS [Desc],
  qryItemMovementPackSales.CartonID,
  Sum(
    CInt(
      ([PacksSold] + 0.01)/ [PacksPerCarton]
    )
  ) AS YTDQty,
  0 AS CurrMM,
  Sum(
    IIf(
      [MM] = "01", [PacksSold] / [PacksPerCarton],
      0
    )
  ) AS Jan,
  Sum(
    IIf(
      [MM] = "02", [PacksSold] / [PacksPerCarton],
      0
    )
  ) AS Feb,
  Sum(
    IIf(
      [MM] = "03", [PacksSold] / [PacksPerCarton],
      0
    )
  ) AS Mar,
  Sum(
    IIf(
      [MM] = "04", [PacksSold] / [PacksPerCarton],
      0
    )
  ) AS Apr,
  Sum(
    IIf(
      [MM] = "05", [PacksSold] / [PacksPerCarton],
      0
    )
  ) AS May,
  Sum(
    IIf(
      [MM] = "06", [PacksSold] / [PacksPerCarton],
      0
    )
  ) AS Jun,
  Sum(
    IIf(
      [MM] = "07", [PacksSold] / [PacksPerCarton],
      0
    )
  ) AS Jul,
  Sum(
    IIf(
      [MM] = "08", [PacksSold] / [PacksPerCarton],
      0
    )
  ) AS Aug,
  Sum(
    IIf(
      [MM] = "09", [PacksSold] / [PacksPerCarton],
      0
    )
  ) AS Sep,
  Sum(
    IIf(
      [MM] = "10", [PacksSold] / [PacksPerCarton],
      0
    )
  ) AS Oct,
  Sum(
    IIf(
      [MM] = "11", [PacksSold] / [PacksPerCarton],
      0
    )
  ) AS Nov,
  Sum(
    IIf(
      [MM] = "12", [PacksSold] / [PacksPerCarton],
      0
    )
  ) AS [Dec]
FROM
  qryItemMovementPackSales
WHERE
  (
    (
      (
        Format([mmyy], "yyyymm")
      )>= Format(
        (
          Now()-365
        ),
        "yyyymm"
      )
      And (
        Format([mmyy], "yyyymm")
      )< Format(
        (
          Now()
        ),
        "yyyymm"
      )
    )
  )
GROUP BY
  " Sales - Packs",
  qryItemMovementPackSales.CartonID,
  0,
  qryItemMovementPackSales.PackID;
