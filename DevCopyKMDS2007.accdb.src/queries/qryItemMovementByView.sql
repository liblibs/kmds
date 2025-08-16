SELECT
  dbo_vwInvoLineByCodeByYearMonth.Code AS Expr1,
  dbo_vwInvoLineByCodeByYearMonth.YR AS YYYY,
  dbo_vwInvoLineByCodeByYearMonth.MM AS Expr2,
  dbo_vwInvoLineByCodeByYearMonth.TransType AS Expr3,
  dbo_vwInvoLineByCodeByYearMonth.Qty AS SumOfQuantity,
  dbo_vwInvoLineByCodeByYearMonth.Sales AS SumOfPrice,
  Format(
    CDate([mm] & "/01/" & [yr]),
    "yyyymm"
  ) AS TranMMYY
FROM
  dbo_vwInvoLineByCodeByYearMonth
WHERE
  (
    (
      (
        Format(
          CDate([mm] & "/01/" & [yr]),
          "yyyymm"
        )
      )>= Format(
        (
          Now()-365
        ),
        "yyyymm"
      )
      And (
        Format(
          CDate([mm] & "/01/" & [yr]),
          "yyyymm"
        )
      )< Format(
        (
          Now()
        ),
        "yyyymm"
      )
    )
  );
