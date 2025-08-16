UPDATE
  MetrologicScanbatch
SET
  MetrologicScanbatch.Quantity = [txtQuantity]
WHERE
  (
    (
      (
        Len([txtQuantity])
      )< 4
    )
  );
