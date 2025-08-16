PARAMETERS parUser Text (255),
parDate DateTime; INSERT INTO pcChargeDebitsAndCredits (
  Station, [Date], [Time], [Number],
  Card, Member, Result, Auth, Amount,
  CapturedAmount, DebitOrCredit, TicketNumber
)
SELECT
  pcChargeTrans.Station AS Expr1,
  pcChargeTrans.Date AS Expr2,
  pcChargeTrans.Time AS Expr3,
  pcChargeTrans.Number AS Expr4,
  pcChargeTrans.Card AS Expr5,
  pcChargeTrans.Member AS Expr6,
  pcChargeTrans.Result AS Expr7,
  pcChargeTrans.Auth AS Expr8,
  pcChargeTrans.Amount AS Expr9,
  IIf(
    [Result] = "CAPTURED", [Total_Auth],
    0
  ) AS CapturedAmount,
  "C" AS DC,
  IIf(
    IsNull([Ticket])
    Or [Ticket] = "",
    0,
    [Ticket]
  ) AS TicketNumber
FROM
  pcChargeTrans
WHERE
  (
    (
      (pcChargeTrans.Station)= [parUser]
    )
    And (
      (pcChargeTrans.Date)= [parDate]
    )
  )
ORDER BY
  pcChargeTrans.Date,
  pcChargeTrans.Time;
