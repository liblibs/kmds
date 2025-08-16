PARAMETERS parTranDate DateTime; INSERT INTO InvoLine5 (
  InvoHeadID, ID, CashSysID, CashierID,
  ShiftID, Code, SalesDept, Quantity,
  Price, InputCode, SalesGroup, Drawer,
  TransType, UpdateOnhand, SetupDate
)
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
  InvoLine.SetupDate
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.TransType)= "S"
      Or (InvoLine.TransType)= "C"
    )
    AND (
      (InvoLine.SetupDate) Between [parTranDate]
      And ([parTranDate] + 1)
    )
  );
