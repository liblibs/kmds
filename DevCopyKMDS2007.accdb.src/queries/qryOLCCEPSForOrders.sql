PARAMETERS prmOrderNumber Long,
prmOrderDate DateTime;
SELECT
  OLCCEPS.OLCCItem,
  Sum(OLCCEPS.Shipped) AS SumOfShipped,
  First(OLCCEPS.Price) AS FirstOfPrice,
  First(
    vwItemsWithCurrentPricing.CurPrice
  ) AS FirstOfCurPrice,
  First(OLCCEPS.UnitsPerCase) AS FirstOfUnitsPerCase,
  OLCCEPS.OrderNumber,
  OLCCEPS.OrderDate
FROM
  OLCCEPS
  INNER JOIN vwItemsWithCurrentPricing ON OLCCEPS.OLCCItem = vwItemsWithCurrentPricing.ID
GROUP BY
  OLCCEPS.OLCCItem,
  OLCCEPS.OrderNumber,
  OLCCEPS.OrderDate
HAVING
  (
    (
      (OLCCEPS.OrderNumber)= [prmOrderNumber]
    )
  )
ORDER BY
  OLCCEPS.OLCCItem;
