UPDATE
  OLCCEPS
  INNER JOIN Item ON OLCCEPS.OLCCItem = Item.ID
SET
  Item.UnitsPerCase = [OLCCEPS].[UnitsPerCase],
  Item.ChangeDate = Now()
WHERE
  (
    (
      (Item.UnitsPerCase)<> [OLCCEPS].[UnitsPerCase]
    )
  );
