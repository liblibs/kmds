INSERT INTO InvoLine (
  SetupDate, Quantity, InvoHeadID, CashSysID,
  CashierID, Code, InputCode, TransType,
  id
)
SELECT
  #8/2/2020# AS sd,
  sheet2x.Aug,
  999 AS InvoHeadID,
  99 AS CashSysID,
  99 AS CashierID,
  sheet2x.ItemID,
  "KMDS Conversion" AS InputCode,
  "S" AS TransType,
  sheet2x.id
FROM
  sheet2x
WHERE
  (
    (
      (sheet2x.Aug)<> 0
    )
  );
