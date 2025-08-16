SELECT
  qryClubOrderGuidesReportAndSysfile.*,
  IIf(
    IsNull([Week1]),
    0,
    [Week1]
  ) AS W1,
  IIf(
    IsNull([Week2]),
    0,
    [Week2]
  ) AS W2,
  IIf(
    IsNull([Week3]),
    0,
    [Week3]
  ) AS W3,
  IIf(
    IsNull([Week4]),
    0,
    [Week4]
  ) AS W4,
  IIf(
    IsNull([Week5]),
    0,
    [Week5]
  ) AS W5,
  IIf(
    ([ItemID] > 99900000000),
    [ItemID] - 99900000000,
    [ItemID]
  ) AS STRIPPED
FROM
  qryClubOrderGuidesReportAndSysfile
  LEFT JOIN qryClubSalesThisMonthByItemByWeek ON (
    qryClubOrderGuidesReportAndSysfile.ItemID = qryClubSalesThisMonthByItemByWeek.Item.ID
  )
  AND (
    qryClubOrderGuidesReportAndSysfile.ID = qryClubSalesThisMonthByItemByWeek.Club.ID
  );
