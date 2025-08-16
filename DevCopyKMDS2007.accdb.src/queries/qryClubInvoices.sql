SELECT
  InvoLineAllRegisters.TransType,
  InvoLineAllRegisters.InvoHeadID,
  InvoLineAllRegisters.SalesDept AS ClubID
FROM
  InvoLineAllRegisters
GROUP BY
  InvoLineAllRegisters.TransType,
  InvoLineAllRegisters.InvoHeadID,
  InvoLineAllRegisters.SalesDept
HAVING
  (
    (
      (InvoLineAllRegisters.TransType)= "C"
    )
  );
