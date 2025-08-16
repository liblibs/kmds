SELECT
  SalesDepartment.ID,
  SalesDepartment.Description
FROM
  SalesDepartment
WHERE
  (
    (
      (SalesDepartment.ID)> 2
    )
  )
ORDER BY
  SalesDepartment.ID;
