UPDATE
  dbo_OLCCitem
  INNER JOIN Item ON dbo_OLCCitem.Code = Item.ID
SET
  Item.UnitsPerCase = [dbo_OLCCitem].[UnitsPerCase]
WHERE
  (
    (
      (Item.UnitsPerCase)<> [dbo_OLCCitem].[UnitsPerCase]
    )
  );
