SELECT
  Cashier.ID,
  Cashier.Password,
  Cashier.Name,
  Cashier.AllowChangePrice,
  Cashier.AllowChangeOnhand,
  Cashier.AllowChangeCost,
  Cashier.AllowChangeOrderAt,
  Cashier.AllowChangeOrderTo,
  Cashier.AllowChangeDescription,
  Cashier.AllowHourlyWageView,
  Cashier.AllowDeleteDocument,
  Cashier.PasswordDate,
  Cashier.AccountLocked,
  Cashier.AccountLockedDate
FROM
  Cashier
WHERE
  (
    (
      (Cashier.ID)= [parID]
    )
    AND (
      (Cashier.Password)= [parPassword]
    )
  );
