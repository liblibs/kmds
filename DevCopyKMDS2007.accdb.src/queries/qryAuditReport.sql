SELECT
  Audit.ItemID,
  Item.Description,
  Size.Description AS SizeDesc,
  Audit.OnShelf,
  Audit.CountFront,
  Audit.CountBack,
  Audit.CountDisplay,
  Audit.CountOverflow,
  Audit.PhysicalCount,
  [PhysicalCount] - [onshelf] AS Var,
  Audit.Price,
  Audit.Cost,
  Audit.SetupDate,
  Audit.ChangeDate,
  ([PhysicalCount] - [onshelf])* Audit.Price AS ExtPrice,
  ([PhysicalCount] - [onshelf])* Audit.Cost AS ExtCost,
  Audit.VarianceAmount
FROM
  (
    Audit
    INNER JOIN Item ON Audit.ItemID = Item.ID
  )
  INNER JOIN [Size] ON Item.SizeID = Size.ID
WHERE
  (
    (
      (
        IIf(
          Forms!Audit!chkOnlyVariance,-1,
          0
        )
      )= 0
    )
  )
  Or (
    (
      (Audit.VarianceAmount)<> 0
    )
    And (
      (
        IIf(
          Forms!Audit!chkOnlyVariance,-1,
          0
        )
      )=-1
    )
  );
