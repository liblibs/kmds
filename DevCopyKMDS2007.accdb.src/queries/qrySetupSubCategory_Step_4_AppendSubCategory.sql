INSERT INTO SubCategory (Description, CategoryID)
SELECT
  OLCCWebItems.Category,
  1 AS CatID
FROM
  OLCCWebItems
GROUP BY
  OLCCWebItems.Category,
  1;
