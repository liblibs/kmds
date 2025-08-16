UPDATE
  (
    qryProductListWithCategoriesFromOLCCWebsite
    INNER JOIN SubCategory ON qryProductListWithCategoriesFromOLCCWebsite.F5 = SubCategory.Description
  )
  INNER JOIN Item ON qryProductListWithCategoriesFromOLCCWebsite.Code = Item.ID
SET
  Item.SubCategoryID = SubCategory.ID;
