SELECT
  Item.SubCategoryID,
  ItemLabels.ItemD,
  ItemLabels.curprice AS Price,
  ItemLabels.ID
FROM
  ItemLabels
  INNER JOIN Item ON ItemLabels.ID = Item.ID
WHERE
  (
    (
      (Item.SubCategoryID) Between [Forms]![LabelList]![cboCategoryFrom]
      And [Forms]![LabelList]![cboCategoryTo]
    )
  )
ORDER BY
  ItemLabels.ItemD;
