SELECT
  InvoLine5.TransType,
  InvoLine5.InvoHeadID,
  InvoLine5.SalesDept AS ClubID
FROM
  InvoLine5
GROUP BY
  InvoLine5.TransType,
  InvoLine5.InvoHeadID,
  InvoLine5.SalesDept
HAVING
  (
    (
      (InvoLine5.TransType)= "C"
    )
  );
