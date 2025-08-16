SELECT
  qryInvoline5Charges.Auth,
  qryInvoline5Charges.InvoHeadID,
  qryInvoline5Charges.CashSysID,
  qryInvoline5Charges.CashierID,
  qryInvoline5Charges.ShiftID,
  qryInvoline5Charges.Code,
  qryInvoline5Charges.Quantity,
  qryInvoline5Charges.Price,
  pcChargeTrans.Amount,
  [Price] - [Amount] AS Diff,
  pcChargeTrans.TroutD
FROM
  qryInvoline5Charges
  INNER JOIN pcChargeTrans ON qryInvoline5Charges.Auth = pcChargeTrans.Auth
WHERE
  (
    (
      ([Price] - [Amount])<> 0
    )
  );
