UPDATE
  Item
  LEFT JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
SET
  Item.SubCategoryID = 0
WHERE
  (
    (
      (SubCategory.ID) Is Null
    )
  );
