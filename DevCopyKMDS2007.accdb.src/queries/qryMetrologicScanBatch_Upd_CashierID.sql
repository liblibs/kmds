PARAMETERS parCashierID Long;
UPDATE
  MetrologicScanbatch
SET
  MetrologicScanbatch.CashierID = [parCashierID]
WHERE
  (
    (
      (MetrologicScanbatch.CashierID)= 0
    )
  );
