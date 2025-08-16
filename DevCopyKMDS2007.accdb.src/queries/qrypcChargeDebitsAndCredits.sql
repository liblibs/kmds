SELECT
  pcChargeDebitsAndCredits.Station,
  pcChargeDebitsAndCredits.Date,
  pcChargeDebitsAndCredits.Time,
  pcChargeDebitsAndCredits.Number,
  pcChargeDebitsAndCredits.Card,
  pcChargeDebitsAndCredits.Member,
  pcChargeDebitsAndCredits.Result,
  pcChargeDebitsAndCredits.Auth,
  pcChargeDebitsAndCredits.Amount,
  pcChargeDebitsAndCredits.CapturedAmount,
  IIf(
    [Result] = "Captured", [DebitOrCredit],
    "-"
  ) AS DC
FROM
  pcChargeDebitsAndCredits
ORDER BY
  pcChargeDebitsAndCredits.Date,
  pcChargeDebitsAndCredits.Time;
