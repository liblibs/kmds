INSERT INTO ITEM (
  ID, Description, Price, UnitsPerCase,
  StockNumber, SizeID, SalesDepartmentID,
  VendorID, Cost
)
SELECT
  dbo_OLCCitem.Code,
  dbo_OLCCitem.Description,
  dbo_OLCCitem.Price,
  dbo_OLCCitem.UnitsPerCase,
  dbo_OLCCitem.StockStatus,
  Right([CODE], 1) AS SD,
  1 AS SLD,
  1 AS VD,
  dbo_OLCCitem.Price
FROM
  WilsonvilleOrderNOF
  INNER JOIN dbo_OLCCitem ON WilsonvilleOrderNOF.ItemID = dbo_OLCCitem.Code;
