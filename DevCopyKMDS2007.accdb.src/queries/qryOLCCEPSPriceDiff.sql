SELECT
  OLCCEPS.OrderNumber,
  vwItemsWithCurrentPricing.ID,
  vwItemsWithCurrentPricing.Description,
  OLCCEPS.Shipped,
  OLCCEPS.Price AS OLAS,
  vwItemsWithCurrentPricing.CurPrice AS KMDS,
  (
    [shipped] *([CURprice] - [olcceps].[Price])
  ) AS Diff
FROM
  OLCCEPS
  INNER JOIN vwItemsWithCurrentPricing ON OLCCEPS.OLCCItem = vwItemsWithCurrentPricing.ID
WHERE
  (
    (
      (OLCCEPS.OrderNumber)= Forms!OLCCPackingList!OrderNumber
    )
    And (
      (OLCCEPS.Price)<> [CurPrice]
    )
  );
