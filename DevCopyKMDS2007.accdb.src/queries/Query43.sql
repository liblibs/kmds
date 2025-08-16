UPDATE
  (
    ScioInitialOrder
    INNER JOIN Item ON ScioInitialOrder.Code = Item.ID
  )
  INNER JOIN SubCategory ON ScioInitialOrder.Category = SubCategory.Description
SET
  Item.SubCategoryID = [SubCategory].[id];
