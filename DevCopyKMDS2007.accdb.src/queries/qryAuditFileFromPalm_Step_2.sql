UPDATE
  PalmFrom
SET
  PalmFrom.ItemID = [Code]
WHERE
  (
    (
      (
        IsNumeric([Code])
      )= True
    )
  );
