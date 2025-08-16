SELECT
  qrySalesByHour.Date,
  qrySalesByHour.SumOfPrice,
  qrySalesByHour.[10 AM],
  qrySalesByHour.[11 AM],
  qrySalesByHour.[12 AM],
  qrySalesByHour.[1 PM],
  qrySalesByHour.[2 PM],
  qrySalesByHour.[3 PM],
  qrySalesByHour.[4 PM],
  qrySalesByHour.[5 PM],
  qrySalesByHour.[6 PM],
  qrySalesByHour.[7 PM],
  qrySalesByHour.[8 PM],
  qrySalesByHour.[9 PM],
  qrySalesByHour.Licensee,
  qrySalesByHour.Count AS LineCount,
  qrySalesByHourInvoiceTotals.CustomerCount AS [Count]
FROM
  qrySalesByHour
  LEFT JOIN qrySalesByHourInvoiceTotals ON qrySalesByHour.Date = qrySalesByHourInvoiceTotals.BusinessDate;
