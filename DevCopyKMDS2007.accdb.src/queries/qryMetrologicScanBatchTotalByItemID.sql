SELECT
  UPCItem.ItemID,
  Sum(MetrologicScanbatch.Quantity) AS SumOfQuantity,
  First(MetrologicScanbatch.CashierID) AS FirstOfCashierID,
  First(MetrologicScanbatch.ID) AS InvoLineID
FROM
  MetrologicScanbatch
  INNER JOIN UPCItem ON MetrologicScanbatch.ScanCode = UPCItem.ID
GROUP BY
  UPCItem.ItemID;
