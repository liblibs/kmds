UPDATE
  MetrologicScanbatch
SET
  MetrologicScanbatch.txtQuantity = "0"
WHERE
  (
    (
      (
        IsNumeric([txtquantity])
      )= False
    )
  );
