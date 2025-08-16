INSERT INTO InvoLine5 (
  SetupDate, InvoHeadID, ID, CashSysID,
  CashierID, ShiftID, Code, SalesDept,
  Quantity, Price, InputCode, SalesGroup,
  Drawer, TransType, UpdateOnhand,
  Cost
)
SELECT
  InvoLine.SetupDate,
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
  InvoLine.Cost
FROM
  InvoLine
WHERE
  (
    (
      (InvoLine.SetupDate) Between [Forms]![PrintInventoryReport]![txtDateRangeFrom]
      And [Forms]![PrintInventoryReport]![txtDateRangeTo]
    )
  );
