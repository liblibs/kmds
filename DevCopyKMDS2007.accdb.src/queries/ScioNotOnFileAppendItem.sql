INSERT INTO Item (
  ID, SizeID, VendorID, Description,
  Price, Cost, UnitsPerCase, UnitsPerOrder,
  SalesDepartmentId
)
SELECT
  ScioNotOnFile.ID,
  ScioNotOnFile.SizeID,
  ScioNotOnFile.VendorID,
  ScioNotOnFile.Description,
  ScioNotOnFile.Price,
  ScioNotOnFile.Cost,
  ScioNotOnFile.UnitsPerCase,
  ScioNotOnFile.UnitsPerOrder,
  ScioNotOnFile.SalesDepartmentId
FROM
  ScioNotOnFile;
