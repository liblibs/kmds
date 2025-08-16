SELECT
  qryPunchDates.PunchDate AS Expr1,
  qryPunchDates.CashierID AS Expr2,
  qryPunchDates.FirstPunch AS Expr3,
  qryPunchDates.LastPunch AS Expr4,
  (
    DateDiff("n", [FirstPunch], [LastPunch])/ 60
  ) AS HoursWorked
FROM
  qryPunchDates;
