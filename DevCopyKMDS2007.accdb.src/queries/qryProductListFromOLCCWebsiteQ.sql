SELECT
  DISTINCTROW CLng(
    Left(
      Trim([f1]),
      Len(
        Trim([f1])
      )-1
    )& l2n(
      Right(
        Trim([F1]),
        1
      )
    )
  ) AS Code,
  Right(
    [f2],
    (
      Len([f2])- InStr([f2], " ")
    )
  ) AS Description,
  CCur([f3]) AS Price,
  CCur(
    Left(
      [f2],
      InStr([f2], " ")
    )
  )/ CCur([f3]) AS CaseQ,
  L2N(
    Right([F1], 1)
  ) AS [Size],
  ProductListFromOLCCWebsite.F1 AS Expr1,
  l2n(
    Right(
      Trim([F1]),
      1
    )
  ) AS ln INTO qryProductListFromOLCCWebsite
FROM
  ProductListFromOLCCWebsite;
