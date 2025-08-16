INSERT INTO InvoLine (
  SetupDate, Quantity, InvoHeadID, ID,
  CashSysID, CashierID, ShiftID, Code,
  SalesDept, Price, TransType
)
SELECT
  "9/2/2011" AS SetupDate,
  KenoSalesHist.sep AS Quantity,
  9999 AS InvoHeadID,
  KenoSalesHist.ID,
  99 AS CashSysID,
  99 AS CashierID,
  99 AS ShiftID,
  KenoSalesHist.itemid AS Code,
  1 AS SalesDept,
  [Item].[price] * [Quantity] AS Prices,
  "S" AS TransType
FROM
  KenoSalesHist
  INNER JOIN Item ON KenoSalesHist.ItemID = Item.ID
WHERE
  (
    (
      (KenoSalesHist.sep)<> 0
    )
  );
