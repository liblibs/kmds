SELECT
  InvoLine.InvoHeadID,
  InvoLine.ID,
  InvoLine.CashSysID,
  InvoLine.CashierID,
  InvoLine.ShiftID,
  InvoLine.Code,
  InvoLine.SalesDept,
  InvoLine.Quantity,
  InvoLine.Price,
  InvoLine.InputCode,
  InvoLine.SalesGroup,
  InvoLine.Drawer,
  InvoLine.TransType,
  InvoLine.UpdateOnhand,
  InvoLine.Cost,
  InvoLine.SetupDate
FROM
  205hEAD
  INNER JOIN InvoLine ON (
    [205hEAD].SetupDate = InvoLine.SetupDate
  )
  AND (
    [205hEAD].CashSysID = InvoLine.CashSysID
  )
  AND (
    [205hEAD].InvoHeadID = InvoLine.InvoHeadID
  )
WHERE
  (
    (
      (InvoLine.TransType)= "S"
    )
  );
