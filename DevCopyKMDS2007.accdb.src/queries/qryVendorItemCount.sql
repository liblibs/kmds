SELECT
  Item.VendorID,
  Sum(1) AS [Count]
FROM
  Item
GROUP BY
  Item.VendorID;
