SELECT
  UPCItem.ItemID,
  Sum(1) AS [Count]
FROM
  UPCItem
GROUP BY
  UPCItem.ItemID;
