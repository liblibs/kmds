SELECT
  Audit.ItemID,
  Audit.OnShelf,
  Audit.CountFront,
  Audit.CountBack,
  Audit.PhysicalCount,
  Audit.CountDisplay,
  Audit.CountOverflow,
  (
    [CountFront] + [CountBack] + [CountDisplay] + [CountOverflow]
  )- [onshelf] AS Var,
  Audit.Price,
  Audit.Cost,
  Audit.SetupDate,
  Audit.ChangeDate,
  ([Var])* [Price] AS ExtPrice,
  ([Var])* [Cost] AS ExtCost,
  Audit.VarianceAmount,
  Audit.ID
FROM
  Audit;
