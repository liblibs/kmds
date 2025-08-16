INSERT INTO SubCategory (Description, CategoryID)
SELECT
  [OLCC Categories].Category,
  1 AS Cat
FROM
  [OLCC Categories]
GROUP BY
  [OLCC Categories].Category,
  1;
