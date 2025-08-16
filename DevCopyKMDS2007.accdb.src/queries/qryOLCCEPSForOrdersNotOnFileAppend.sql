PARAMETERS prmOrderNumber Long,
prmOrderDate DateTime; INSERT INTO item (
  ID, SizeID, SalesDepartmentID, VendorID,
  Description, UnitsPerCase, Price,
  Cost
)
SELECT
  OLCCEPS.OLCCItem,
  Right([OLCCItem], 2) AS SizeId,
  1 AS SalesDepartmentID,
  1 AS VendorID,
  OLCCEPS.Description,
  OLCCEPS.UnitsPerCase,
  OLCCEPS.Price,
  OLCCEPS.price
FROM
  OLCCEPS
  LEFT JOIN Item ON OLCCEPS.OLCCItem = Item.ID
WHERE
  (
    (
      (OLCCEPS.OrderNumber)= [prmOrderNumber]
    )
    AND (
      (Item.ID) Is Null
    )
    AND (
      (OLCCEPS.OrderDate)= [prmOrderDate]
    )
  )
ORDER BY
  OLCCEPS.OLCCItem;
