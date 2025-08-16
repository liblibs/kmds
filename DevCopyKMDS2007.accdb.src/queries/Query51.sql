INSERT INTO InvoLine (
  InvoHeadID, SetupDate, Quantity, ID,
  Code, InputCode, TransType
)
SELECT
  99901 AS InvoHeadID,
  #11/1/2019# AS SetupDate,
  Sheet1.NOV AS Quantity,
  Sheet1.ID,
  Sheet1.ITEMID AS Code,
  "KMDS Conv" AS InputCode,
  "S" AS TransType
FROM
  Sheet1,
  InvoLine;
