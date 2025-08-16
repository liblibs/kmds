INSERT INTO InvoLine (
  InvoHeadID, ID, CashSysID, Code, Quantity,
  InputCode, TransType, SetupDate
)
SELECT
  99 AS InvoHeadID,
  GaribaldiSalesHist.ID,
  99 AS CashSysID,
  GaribaldiSalesHist.itemid AS Code,
  GaribaldiSalesHist.Oct AS Quantity,
  "Conversion" AS InputCode,
  "S" AS TransType,
  #10/2/2014# AS SetupDate
FROM
  GaribaldiSalesHist
WHERE
  (
    (
      (GaribaldiSalesHist.Oct)<> 0
    )
  );
