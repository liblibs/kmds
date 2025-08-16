UPDATE
  ItemLabels
SET
  ItemLabels.Labels = 0
WHERE
  (
    (
      (ItemLabels.Labels)<> 0
    )
  );
