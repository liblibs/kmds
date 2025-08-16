SELECT
  SubCategory.Description,
  Item.ID,
  Item.Description,
  Size.Description AS SizeDesc,
  Size.CaseQty,
  Item.Price,
  Max(UPCItem.ID) AS UPCCode
FROM
  (
    UPCItem
    INNER JOIN (
      Item
      INNER JOIN SubCategory ON Item.SubCategoryID = SubCategory.ID
    ) ON UPCItem.ItemID = Item.ID
  )
  INNER JOIN [Size] ON Item.SizeID = Size.ID
WHERE
  (
    (
      (
        Len([UPCItem].[ID])
      )> 7
    )
  )
GROUP BY
  SubCategory.Description,
  Item.ID,
  Item.Description,
  Size.Description,
  Size.CaseQty,
  Item.Price
ORDER BY
  SubCategory.Description,
  Item.ID;
