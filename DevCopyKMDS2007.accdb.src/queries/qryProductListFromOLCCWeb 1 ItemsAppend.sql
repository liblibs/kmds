INSERT INTO item (
  ID, SizeID, SalesDepartmentID, VendorID,
  Description, Price, UnitsPerCase,
  UnitsPerOrder
)
SELECT
  qryProductListFromOLCCWebsite.Code,
  qryProductListFromOLCCWebsite.Size,
  1 AS sd,
  1 AS VendorID,
  qryProductListFromOLCCWebsite.Description,
  qryProductListFromOLCCWebsite.Price,
  qryProductListFromOLCCWebsite.CaseQ,
  qryProductListFromOLCCWebsite.CaseQ
FROM
  Item
  RIGHT JOIN qryProductListFromOLCCWebsite ON Item.ID = qryProductListFromOLCCWebsite.Code
WHERE
  (
    (
      (Item.ID) Is Null
    )
  );
