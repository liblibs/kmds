INSERT INTO InvoLine (
  SetupDate, Quantity, InvoHeadID, CashSysID,
  CashierID, Code, SalesDept, InputCode,
  TransType, ID
)
SELECT
  #4/1/2018# AS SetupDate,
  EugSouthHill.Apr,
  9999 AS InvoHeadID,
  99 AS CashSysID,
  99 AS CashierID,
  EugSouthHill.Code,
  1 AS SalesDept,
  "KMDS Convert" AS InputCode,
  "S" AS TransType,
  EugSouthHill.ID
FROM
  EugSouthHill
WHERE
  (
    (
      (EugSouthHill.Apr)<> 0
    )
  );
