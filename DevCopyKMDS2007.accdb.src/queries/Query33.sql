SELECT
  Item.ID,
  Item.SalesDepartmentID,
  Item.Description,
  Item.SizeID,
  Item.SetupDate
FROM
  Item
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
    )
    AND (
      (Item.SetupDate)> #7/27/2015#
    )
  );
