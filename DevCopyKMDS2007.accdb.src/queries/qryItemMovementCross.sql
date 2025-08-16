TRANSFORM Sum(qryItemMovement.SumOfQuantity) AS SumOfSumOfQuantity
SELECT
  qryItemMovement.Code,
  qryItemMovement.TransType
FROM
  qryItemMovement
GROUP BY
  qryItemMovement.Code,
  qryItemMovement.TransType PIVOT qryItemMovement.MM;
