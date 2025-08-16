SELECT
  Cashier.ID,
  Cashier.Password,
  Cashier.PasswordHistory1,
  Cashier.PasswordHistory2,
  Cashier.PasswordHistory3,
  Cashier.PasswordHistory4
FROM
  Cashier
WHERE
  (
    (
      (Cashier.ID)= [parCashierID]
    )
  );
