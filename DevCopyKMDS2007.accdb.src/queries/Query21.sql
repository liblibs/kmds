SELECT
  dbo_OLCCitem.Code,
  dbo_OLCCitem.Description,
  dbo_OLCCitem.Price,
  dbo_OLCCitem.UnitsPerCase,
  dbo_OLCCitem.StockStatus,
  Right([Code], 1) AS SizeID
FROM
  dbo_OLCCitem
  LEFT JOIN Item ON dbo_OLCCitem.Code = Item.ID
WHERE
  (
    (
      (Item.ID) Is Null
    )
  );
