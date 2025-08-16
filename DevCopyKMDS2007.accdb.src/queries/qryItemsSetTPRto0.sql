UPDATE
  Item
SET
  Item.TPR = 0
WHERE
  (
    (
      (Item.TPR)<> 0
    )
  );
