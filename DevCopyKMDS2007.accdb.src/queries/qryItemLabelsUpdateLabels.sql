PARAMETERS parItemID Long;
UPDATE
  ItemLabels
SET
  ItemLabels.Labels = 1
WHERE
  (
    (
      (ItemLabels.ID)= [parItemID]
    )
  );
