PARAMETERS parCashierID Long;
SELECT
  CashiersAndPrograms.CashierID,
  CashiersAndPrograms.Allowed,
  Sum(1) AS [Count]
FROM
  CashiersAndPrograms
GROUP BY
  CashiersAndPrograms.CashierID,
  CashiersAndPrograms.Allowed
HAVING
  (
    (
      (CashiersAndPrograms.CashierID)= [parCashierID]
    )
    AND (
      (CashiersAndPrograms.Allowed)<> 0
    )
  );
