UPDATE
  InvoLine4
SET
  InvoLine4.Drawer = Right([code], 1)
WHERE
  (
    (
      (
        IsNumeric([code])
      )= True
    )
  );
