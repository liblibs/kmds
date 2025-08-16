SELECT
  OLCCEPS.OLCCItem,
  OLCCEPS.Description,
  OLCCEPS.UnitsPerCase,
  OLCCEPS.Shipped,
  OLCCEPS.Price,
  OLCCEPS.UnitsPerCase,
  OLCCEPS.OrderNumber,
  Item.ID,
  OLCCEPS.OrderNumber
FROM
  OLCCEPS
  LEFT JOIN Item ON OLCCEPS.OLCCItem = Item.ID
WHERE
  (
    (
      (Item.ID) Is Null
    )
    And (
      (OLCCEPS.OrderNumber)= Forms!OLCCPackingList!OrderNumber
    )
  )
ORDER BY
  OLCCEPS.OLCCItem;
