SELECT
  v.SalesDptDescr,
  v.CategoryDescr AS Category,
  i.ID,
  v.Description AS Description,
  s.Description AS [Size],
  i.RestockAmount,
  [RestockAmount] * v.Price AS Dollars,
  v.OnShelf,
  i.Location,
  UCASE(d.[Description]) AS SalesDepartment
FROM
  (
    (
      vwItemsWithCurrentPricing AS v
      INNER JOIN Item AS i ON v.ID = i.ID
    )
    INNER JOIN [Size] AS s ON i.SizeID = s.ID
  )
  INNER JOIN SalesDepartment AS d ON v.SalesDepartmentID = d.ID
WHERE
  i.ID Is Not Null
  And i.RestockAmount <> 0
ORDER BY
  v.CategoryDescr,
  s.Description,
  i.ID;
