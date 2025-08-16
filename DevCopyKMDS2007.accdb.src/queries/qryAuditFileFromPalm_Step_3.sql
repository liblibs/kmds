UPDATE
  PalmFrom
  INNER JOIN Audit ON PalmFrom.ItemID = Audit.ItemID
SET
  Audit.PhysicalCount = PalmFrom.Physical,
  Audit.VarianceAmount = PalmFrom.Physical - Audit.OnShelf,
  Audit.CountFront = PalmFrom.Physical,
  Audit.CountBack = 0,
  Audit.CountDisplay = 0,
  Audit.CountOverflow = 0
WHERE
  (
    (
      (PalmFrom.Physical) Is Not Null
    )
  );
