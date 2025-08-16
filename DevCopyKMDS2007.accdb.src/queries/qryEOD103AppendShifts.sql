INSERT INTO Shifts (ID, CashSysID, BusinessDate)
SELECT
  InvoHead5.ShiftID,
  InvoHead5.CashSysID,
  Last(
    Format(
      InvoHead5.SetupDate, "mm/dd/yyyy"
    )
  ) AS BusinessDate
FROM
  InvoHead5
  LEFT JOIN Shifts ON (InvoHead5.ShiftID = Shifts.ID)
  AND (
    InvoHead5.CashSysID = Shifts.CashSysID
  )
WHERE
  (
    (
      (Shifts.ID) Is Null
    )
  )
GROUP BY
  InvoHead5.ShiftID,
  InvoHead5.CashSysID
ORDER BY
  InvoHead5.ShiftID,
  InvoHead5.CashSysID,
  Last(
    Format(
      InvoHead5.SetupDate, "mm/dd/yyyy"
    )
  );
