INSERT INTO InvoLine (
  InvoHeadID, ID, CashSysID, CashierID,
  ShiftID, Code, SalesDept, Quantity,
  Price, InputCode, SalesGroup, Drawer,
  TransType, UpdateOnhand, SetupDate
)
SELECT
  InvoLine5.InvoHeadID,
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
  InvoLine5;
