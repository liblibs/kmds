SELECT
  CashSys.ID,
  CashSys.ComputerName,
  CashSys.XfrSmode
FROM
  CashSys
WHERE
  (
    (
      (CashSys.ComputerName)<> ""
    )
    AND (
      (CashSys.XfrSmode)<> 0
    )
  )
ORDER BY
  CashSys.ID;
