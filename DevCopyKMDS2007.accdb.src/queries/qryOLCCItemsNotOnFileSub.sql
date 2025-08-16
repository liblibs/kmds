SELECT
  dbo_OLCCitem.Code,
  dbo_OLCCitem.Description,
  dbo_OLCCitem.Price,
  dbo_OLCCitem.UnitsPerCase,
  [OLCC Categories].Category,
  dbo_OLCCitem.StockStatus,
  Right([dbo_OLCCitem.Code], 2) AS SizeID,
  dbo_OLCCitem.SetupDate,
  qrydbo_OLCCScanCode_12Or8.UPCCodeKMDS,
  SubCategory.ID AS SubCategoryID
FROM
  (
    (
      (
        dbo_OLCCitem
        LEFT JOIN Item ON dbo_OLCCitem.Code = Item.ID
      )
      LEFT JOIN [OLCC Categories] ON dbo_OLCCitem.Code = [OLCC Categories].Code
    )
    LEFT JOIN qrydbo_OLCCScanCode_12Or8 ON dbo_OLCCitem.Code = qrydbo_OLCCScanCode_12Or8.Code
  )
  LEFT JOIN SubCategory ON [OLCC Categories].Category = SubCategory.Description
WHERE
  (
    (
      (Item.ID) Is Null
    )
  );
