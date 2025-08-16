INSERT INTO InvoLine5 (
  InvoHeadID, ID, CashSysID, CashierID,
  Code, Price, SetupDate, TransType,
  InputCode, Quantity
)
SELECT
  InvoLine5.Code,
  [CashierID] + 100 AS NewID,
  InvoLine5.CashSysID,
  InvoLine5.CashierID,
  InvoLine5.Code,
  Sum(0 - [Price]) AS Expr1,
  Last(InvoLine5.SetupDate) AS LastOfSetupDate,
  InvoLine5.TransType,
  "Change" AS InputCod,
  1 AS Nquantity
FROM
  InvoLine5,
  Sysfile
WHERE
  (
    (
      (Sysfile.AutoDropCharges)= True
    )
    AND (
      (InvoLine5.InputCode)<> "Drop AutoDrop"
    )
  )
GROUP BY
  [CashierID] + 100,
  InvoLine5.CashSysID,
  InvoLine5.CashierID,
  InvoLine5.Code,
  InvoLine5.TransType,
  "Change",
  1,
  InvoLine5.Code
HAVING
  (
    (
      (
        Sum(0 - [Price])
      )<> 0
    )
    AND (
      (InvoLine5.TransType)= "T"
    )
    AND (
      (InvoLine5.Code)= 3
    )
  );
