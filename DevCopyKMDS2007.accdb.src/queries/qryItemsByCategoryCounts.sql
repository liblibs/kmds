SELECT
  Item.SubCategoryID,
  Sum(1) AS [count]
FROM
  Item
GROUP BY
  Item.SubCategoryID;
