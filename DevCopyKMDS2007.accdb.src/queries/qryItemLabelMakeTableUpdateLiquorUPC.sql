UPDATE
  ItemLabels
SET
  ItemLabels.UPCCodeMax = Format([id], "000000")
WHERE
  (
    (
      (ItemLabels.ID) Between 1000
      And 99999
    )
  );
