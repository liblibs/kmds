SELECT
  qryXCustomerCount.CashSysID,
  qryXCustomerCount.CashierID,
  Sum(1) AS [COUNT],
  Sum(qryXCustomerCount.Sales) AS CashierSales
FROM
  qryXCustomerCount
GROUP BY
  qryXCustomerCount.CashSysID,
  qryXCustomerCount.CashierID;
