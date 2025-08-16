UPDATE
  PalmOrderGuideFrom
SET
  PalmOrderGuideFrom.ItemID = [Code]
WHERE
  (
    (
      (
        IsNumeric([Code])
      )= True
    )
  );
