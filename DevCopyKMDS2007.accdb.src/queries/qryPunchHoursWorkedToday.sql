SELECT
  qryPunchHoursWorked.PunchDate AS Expr1,
  qryPunchHoursWorked.CashierID AS Expr2,
  qryPunchHoursWorked.FirstPunch AS Expr3,
  qryPunchHoursWorked.LastPunch AS Expr4,
  qryPunchHoursWorked.HoursWorked AS Expr5
FROM
  qryPunchHoursWorked
WHERE
  (
    (
      (
        [qryPunchHoursWorked].[PunchDate]
      )= Format(
        Now(),
        "mm/dd/yyyy"
      )
    )
  );
