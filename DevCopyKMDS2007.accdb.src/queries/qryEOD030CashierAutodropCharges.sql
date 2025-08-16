INSERT INTO InvoLine5 (
  InvoHeadID, ID, CashSysID, CashierID,
  Code, Price, SetupDate, TransType,
  InputCode, Quantity
)
SELECT
  InvoLine5.Code,
  InvoLine5.CashierID,
  InvoLine5.CashSysID,
  InvoLine5.CashierID,
  InvoLine5.Code,
  Sum(InvoLine5.Price) AS SumOfPrice,
  Last(InvoLine5.SetupDate) AS LastOfSetupDate,
  InvoLine5.TransType,
  "Drop AutoDrop" AS InputCod,
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
  InvoLine5.CashSysID,
  InvoLine5.CashierID,
  InvoLine5.Code,
  InvoLine5.TransType,
  "Drop AutoDrop",
  1,
  InvoLine5.CashierID,
  InvoLine5.Code
HAVING
  (
    (
      (
        Sum(InvoLine5.Price)
      )<> 0
    )
    AND (
      (InvoLine5.TransType)= "T"
    )
    AND (
      (InvoLine5.Code)= 3
    )
  );
