PARAMETERS prmSalesdepartmentIdBegin Long,
prmSalesdepartmentIdEnd Long,
prmCategoryDescBegin Text (255),
prmCategoryDescEnd Text (255);
SELECT
  MetrologicLookup.UPC,
  MetrologicLookup.Descr,
  MetrologicLookup.Item,
  MetrologicLookup.PRI,
  MetrologicLookup.curPrice,
  MetrologicLookup.Shelf,
  MetrologicLookup.Description,
  Format([UnitsPerCase], "@@@") AS UCase
FROM
  MetrologicLookup
ORDER BY
  MetrologicLookup.UPC;
