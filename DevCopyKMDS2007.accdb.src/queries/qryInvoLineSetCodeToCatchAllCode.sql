PARAMETERS parCode Decimal,
parCodeCatchAll Decimal;
UPDATE
  InvoLine
SET
  InvoLine.Code = [parCodeCatchAll]
WHERE
  (
    (
      (InvoLine.Code)= [parCode]
    )
  );
