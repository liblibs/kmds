SELECT
  CLng([F2]) AS CaseQty,
  CLng(
    Left(
      [f1],
      Len([f1])-1
    )& l2n(
      Right([f1], 1)
    )
  ) AS ID
FROM
  OCTOB;
