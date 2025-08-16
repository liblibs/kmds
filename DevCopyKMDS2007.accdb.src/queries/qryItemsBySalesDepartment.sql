SELECT
  SalesDepartment.ID,
  SalesDepartment.Description,
  qryItemsBySalesDepartmentTotals.Count
FROM
  SalesDepartment
  INNER JOIN qryItemsBySalesDepartmentTotals ON SalesDepartment.ID = qryItemsBySalesDepartmentTotals.SalesDepartmentID
WHERE
  (
    (
      (SalesDepartment.ID)> 2
    )
  )
ORDER BY
  SalesDepartment.Description;
