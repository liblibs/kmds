UPDATE
  MetrologicScanbatch
SET
  MetrologicScanbatch.Quantity = Mid([txtQuantity], 9, 3)
WHERE
  (
    (
      (
        Len([txtQuantity])
      )= 12
    )
    AND (
      (
        MetrologicScanbatch.txtQuantity
      )< "000000001000"
    )
  );
