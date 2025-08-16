PARAMETERS parCashierID Long,
parAllowed Bit;
UPDATE
  CashiersAndPrograms
SET
  CashiersAndPrograms.Allowed = [parAllowed]
WHERE
  (
    (
      (CashiersAndPrograms.CashierID)= [parCashierID]
    )
  );
