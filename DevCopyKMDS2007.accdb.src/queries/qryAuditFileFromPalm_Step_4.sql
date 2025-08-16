UPDATE
  PalmFrom
  INNER JOIN Audit ON PalmFrom.ItemID = Audit.ItemID
SET
  Audit.PhysicalCount = 0,
  Audit.VarianceAmount = 0 - Audit.OnShelf,
  Audit.CountFront = 0,
  Audit.CountBack = 0,
  Audit.CountDisplay = 0,
  Audit.CountOverflow = 0
WHERE
  (
    (
      (PalmFrom.Physical) Is Null
    )
  );
