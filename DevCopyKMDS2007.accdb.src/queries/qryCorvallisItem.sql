INSERT INTO item (
  ID, Description, StockNumber, SalesDepartmentID,
  Location, VendorID, Onhand, OrderAt,
  OrderTo, Price, Cost, UnitsPerCase,
  UnitsPerOrder, DateLastSold, SetupDate,
  ChangeDate
)
SELECT
  CLng(ItemCorv1.sku) AS ID,
  Trim(ItemCorv1.description) AS descrip,
  Trim([mfg part]) AS StockNumber,
  CInt([dept]) AS SalesDeptID,
  ItemCorv1.location AS Expr1,
  CLng([vendor number]) AS VendorID,
  ItemCorv1.OnHand AS Expr2,
  ItemCorv1.OrderAt AS Expr3,
  ItemCorv1.OrderTo AS Expr4,
  ItemCorv1.Price AS Expr5,
  ItemCorv1.Cost AS Expr6,
  ItemCorv1.[case q] AS UnitsPerCase,
  ItemCorv1.[case q] AS UnitsPerOrder,
  ItemCorv2.DateLastSold AS Expr7,
  4 / 1 / 2005 AS dt,
  4 / 1 / 2005 AS dt2
FROM
  ItemCorv1,
  ItemCorv2
WHERE
  (
    (
      (
        Trim([ItemCorv1].[description])
      ) Is Not Null
    )
  );
