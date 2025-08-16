SELECT
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
  )+ 1 AS Week,
  InvoLine.Code,
  Sum(InvoLine.Quantity) AS SumOfQuantity
FROM
  InvoLine
WHERE
  (
    (
      (
        Format([SetupDate], "yyyymm")
      )= Format(
        Now(),
        "yyyymm"
      )
    )
    AND (
      (InvoLine.TransType)= "C"
    )
  )
GROUP BY
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
  InvoLine.Code
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
