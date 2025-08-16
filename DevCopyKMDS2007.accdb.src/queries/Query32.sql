INSERT INTO Item (
  ID, Description, SizeID, SalesDepartmentID,
  VendorID, Price, UnitsPerCase
)
SELECT
  GaribaldiSalesHist.ItemId,
  dbo_OLCCitem.Description,
  Right([itemid], 1) AS SizeID,
  1 AS SalesDepartmentID,
  1 AS VendorID,
  dbo_OLCCitem.Price,
  dbo_OLCCitem.UnitsPerCase
FROM
  (
    GaribaldiSalesHist
    LEFT JOIN Item ON GaribaldiSalesHist.ItemId = Item.ID
  )
  INNER JOIN dbo_OLCCitem ON GaribaldiSalesHist.ItemId = dbo_OLCCitem.Code
WHERE
  (
    (
      (Item.ID) Is Null
    )
  );
