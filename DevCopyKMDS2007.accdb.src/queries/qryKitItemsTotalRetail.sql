SELECT
  KitItem.KitID,
  Sum(KitItem.Price) AS Retail
FROM
  KitItem
GROUP BY
  KitItem.KitID;
