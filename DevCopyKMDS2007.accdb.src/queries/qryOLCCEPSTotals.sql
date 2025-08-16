SELECT
  OLCCEPS.OrderDate,
  OLCCEPS.OrderNumber,
  Sum(OLCCEPS.Shipped) AS SumOfShipped,
  Sum([Shipped] * [Price]) AS Extended
FROM
  OLCCEPS
GROUP BY
  OLCCEPS.OrderDate,
  OLCCEPS.OrderNumber
ORDER BY
  OLCCEPS.OrderDate DESC,
  OLCCEPS.OrderNumber DESC;
