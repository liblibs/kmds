SELECT
  qrySalesByHourInvoices.BusinessDate,
  Sum(1) AS CustomerCount
FROM
  qrySalesByHourInvoices
GROUP BY
  qrySalesByHourInvoices.BusinessDate;
