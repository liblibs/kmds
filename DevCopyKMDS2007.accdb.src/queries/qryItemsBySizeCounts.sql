SELECT
  Item.SizeID,
  Sum(1) AS [count]
FROM
  Item
GROUP BY
  Item.SizeID;
