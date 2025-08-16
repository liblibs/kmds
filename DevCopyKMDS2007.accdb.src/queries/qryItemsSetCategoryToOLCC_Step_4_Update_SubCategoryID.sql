UPDATE
  Item
  INNER JOIN (
    [OLCC Categories]
    INNER JOIN SubCategory ON [OLCC Categories].Category = SubCategory.Description
  ) ON Item.ID = [OLCC Categories].Code
SET
  Item.SubCategoryID = [SubCategory].[ID];
