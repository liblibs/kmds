PARAMETERS prmSalesdepartmentIdBegin Long,
prmSalesdepartmentIdEnd Long,
prmCategoryDescBegin Text (255),
prmCategoryDescEnd Text (255);
SELECT
  Format(
    [UPCItem].[ID], "!@@@@@@@@@@@@@@"
  ) AS UPC,
  Format(
    Left([Description], 9),
    "!@@@@@@@@@"
  ) AS Descr,
  Format([ItemiD], "!@@@@@@@@@@@") AS Item,
  Format(
    Format([Price], "Fixed"),
    "@@@@@@"
  ) AS PRI,
  vwItemsWithCurrentPricing.curPrice,
  vwItemsWithCurrentPricing.SalesDepartmentID,
  vwItemsWithCurrentPricing.CategoryDescr AS CatDesc,
  vwItemsWithCurrentPricing.onshelf AS Shelf,
  vwItemsWithCurrentPricing.Description,
  vwItemsWithCurrentPricing.UnitsPerCase,
  Format([UnitsPerCase], "@@@") AS UCase
FROM
  UPCItem
  INNER JOIN vwItemsWithCurrentPricing ON UPCItem.ItemID = vwItemsWithCurrentPricing.ID
WHERE
  (
    (
      (
        Format(UPCItem.ID, "!@@@@@@@@@@@@@@")
      )> "0001"
    )
    And (
      (
        vwItemsWithCurrentPricing.SalesDepartmentID
      ) Between [prmSalesdepartmentIdBegin]
      And [prmSalesdepartmentIdEnd]
    )
    And (
      (
        vwItemsWithCurrentPricing.CategoryDescr
      ) Between [prmCategoryDescBegin]
      And [prmCategoryDescEnd]
    )
    And (
      (
        Len(UPCItem.id)
      )< 15
    )
  )
ORDER BY
  Format(
    [UPCItem].[ID], "!@@@@@@@@@@@@@@"
  );
