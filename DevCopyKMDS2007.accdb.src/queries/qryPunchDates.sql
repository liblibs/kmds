SELECT
  Format([PunchTime], 'mm/dd/yyyy') AS PunchDate,
  TimeClockDetail.CashierID AS Expr1,
  Min(TimeClockDetail.PunchTime) AS FirstPunch,
  Max(TimeClockDetail.PunchTime) AS LastPunch
FROM
  TimeClockDetail
GROUP BY
  Format([PunchTime], 'mm/dd/yyyy'),
  TimeClockDetail.CashierID;
