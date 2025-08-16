UPDATE
  (
    OLCCWebItems
    INNER JOIN SubCategory ON OLCCWebItems.Category = SubCategory.Description
  )
  INNER JOIN Item ON OLCCWebItems.Code = Item.ID
SET
  Item.SubCategoryID = SubCategory.id;
