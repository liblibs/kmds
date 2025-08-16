PARAMETERS parUser Text (255),
parDate DateTime; INSERT INTO pcChargeDebitsAndCredits (
  Station, [Date], [Time], [Number],
  Card, Member, Result, Auth, Amount,
  CapturedAmount, DebitOrCredit, TicketNumber
)
SELECT
  pcChargeDebit.Station AS Expr1,
  pcChargeDebit.Date AS Expr2,
  pcChargeDebit.Time AS Expr3,
  pcChargeDebit.Number AS Expr4,
  pcChargeDebit.Card AS Expr5,
  pcChargeDebit.Member AS Expr6,
  pcChargeDebit.Result AS Expr7,
  pcChargeDebit.Auth AS Expr8,
  pcChargeDebit.Amount AS Expr9,
  IIf([Result] = "CAPTURED", [Amount], 0) AS CapturedAmount,
  "D" AS DC,
  IIf(
    IsNull([Ticket])
    Or [Ticket] = "",
    0,
    [Ticket]
  ) AS TicketNumber
FROM
  pcChargeDebit
WHERE
  (
    (
      (pcChargeDebit.Station)= [parUser]
    )
    And (
      (pcChargeDebit.Date)= [parDate]
    )
  )
ORDER BY
  pcChargeDebit.Date,
  pcChargeDebit.Time;
