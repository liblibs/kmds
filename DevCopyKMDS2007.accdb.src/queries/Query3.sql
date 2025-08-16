SELECT
  Format([SetupDate], "yyyymm") AS YearMonth,
  InvoLine.SalesDept AS ClubId,
  (
    Format([SetupDate], "ww")- Format(
      (
        Format(
          Now(),
          "mm"
        )& "/01/" & Format(
          Now(),
          "yyyy"
        )
      ),
      "ww"
    )
  )+ 1 AS Weex,
  Format(
    (
      Format(
        Now(),
        "mm"
      )& "/01/" & Format(
        Now(),
        "yyyy"
      )
    ),
    "ww"
  ) AS Week1,
  InvoLine.Code,
  Sum(InvoLine.Quantity) AS SumOfQuantity,
  Sum(
    IIf(
      Format([SetupDate], "w")= 1,
      [quantity],
      0
    )
  ) AS Week,
  InvoLine.SetupDate
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.TransType)= "C"
    )
  )
GROUP BY
  Format([SetupDate], "yyyymm"),
  InvoLine.SalesDept,
  (
    Format([SetupDate], "ww")- Format(
      (
        Format(
          Now(),
          "mm"
        )& "/01/" & Format(
          Now(),
          "yyyy"
        )
      ),
      "ww"
    )
  )+ 1,
  Format(
    (
      Format(
        Now(),
        "mm"
      )& "/01/" & Format(
        Now(),
        "yyyy"
      )
    ),
    "ww"
  ),
  InvoLine.Code,
  InvoLine.SetupDate
HAVING
  (
    (
      (
        Format([SetupDate], "yyyymm")
      )= Format(
        Now(),
        "yyyymm"
      )
    )
  )
ORDER BY
  InvoLine.SalesDept,
  (
    Format([SetupDate], "ww")- Format(
      (
        Format(
          Now(),
          "mm"
        )& "/01/" & Format(
          Now(),
          "yyyy"
        )
      ),
      "ww"
    )
  )+ 1;
