SELECT
  InvoLine5.InvoHeadID,
  Right(
    [InputCode],
    Len([InputCode])-5
  ) AS Auth,
  InvoLine5.ID,
  InvoLine5.CashSysID,
  InvoLine5.CashierID,
  InvoLine5.ShiftID,
  InvoLine5.Code,
  InvoLine5.SalesDept,
  InvoLine5.Quantity,
  InvoLine5.Price,
  InvoLine5.InputCode,
  InvoLine5.SalesGroup,
  InvoLine5.Drawer,
  InvoLine5.TransType,
  InvoLine5.UpdateOnhand,
  InvoLine5.SetupDate
FROM
  InvoLine5
WHERE
  (
    (
      (InvoLine5.Code)= 3
    )
    AND (
      (InvoLine5.TransType)= "T"
    )
  );
