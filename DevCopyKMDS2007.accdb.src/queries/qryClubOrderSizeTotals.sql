SELECT
  qryClubOrderSizeTot.Description,
  qryClubOrderSizeTot.Count,
  qryClubOrderSizeTot.CQ,
  qryClubOrderSizeTot.QS,
  [QS] \[CQ] AS [Full],
  [QS] -(
    ([QS] \[CQ])* [CQ]
  ) AS Broken
FROM
  qryClubOrderSizeTot;
