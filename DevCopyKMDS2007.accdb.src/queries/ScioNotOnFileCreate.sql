SELECT
  ScioInitialOrder.Code AS ID,
  Right([ScioInitialOrder].[code], 1) AS SizeID,
  1 AS VendorID,
  OLCCitem.Description,
  ScioInitialOrder.Price,
  ScioInitialOrder.Price AS Cost,
  OLCCitem.UnitsPerCase,
  OLCCitem.UnitsPerCase AS UnitsPerOrder,
  1 AS SalesDepartmentId INTO ScioNotOnFile
FROM
  (
    ScioInitialOrder
    LEFT JOIN Item ON ScioInitialOrder.Code = Item.ID
  )
  INNER JOIN OLCCitem ON ScioInitialOrder.Code = OLCCitem.Code
WHERE
  (
    (
      (Item.Description) Is Null
    )
  );
