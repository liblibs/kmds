SELECT
  ItemMovement.Code,
  Sum(
    IIf(
      [Description] = "Transfers", 0 - [CurrMM],
      [CurrMM]
    )
  ) AS CurrMMQty,
  Sum(
    IIf(
      [Description] = "Transfers", 0 - [Jan],
      [Jan]
    )
  ) AS JanQty,
  Sum(
    IIf(
      [Description] = "Transfers", 0 - [Feb],
      [Feb]
    )
  ) AS FebQty,
  Sum(
    IIf(
      [Description] = "Transfers", 0 - [Mar],
      [Mar]
    )
  ) AS MarQty,
  Sum(
    IIf(
      [Description] = "Transfers", 0 - [Apr],
      [Apr]
    )
  ) AS AprQty,
  Sum(
    IIf(
      [Description] = "Transfers", 0 - [May],
      [May]
    )
  ) AS MayQty,
  Sum(
    IIf(
      [Description] = "Transfers", 0 - [Jun],
      [Jun]
    )
  ) AS JunQty,
  Sum(
    IIf(
      [Description] = "Transfers", 0 - [Jul],
      [Jul]
    )
  ) AS JulQty,
  Sum(
    IIf(
      [Description] = "Transfers", 0 - [Aug],
      [Aug]
    )
  ) AS AugQty,
  Sum(
    IIf(
      [Description] = "Transfers", 0 - [Sep],
      [Sep]
    )
  ) AS SepQty,
  Sum(
    IIf(
      [Description] = "Transfers", 0 - [Oct],
      [Oct]
    )
  ) AS OctQty,
  Sum(
    IIf(
      [Description] = "Transfers", 0 - [Nov],
      [Nov]
    )
  ) AS NovQty,
  Sum(
    IIf(
      [Description] = "Transfers", 0 - [Dec],
      [Dec]
    )
  ) AS DecQty,
  Sum(
    IIf(
      [Description] = "Transfers", 0 - [YTDQty],
      [YTDQty]
    )
  ) AS YtdQ
FROM
  ItemMovement
WHERE
  (
    (
      (ItemMovement.Description)= " Sales - Counter"
      Or (ItemMovement.Description)= " Sales - Clubs"
      Or (ItemMovement.Description)= "Transfers"
    )
  )
GROUP BY
  ItemMovement.Code
ORDER BY
  ItemMovement.Code;
