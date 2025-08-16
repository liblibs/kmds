SELECT
  Item.SalesDepartmentID,
  Item.ID,
  Item.SalesDepartmentID,
  Item.SizeID,
  Item.Description,
  Item.Onhand
FROM
  Item
  LEFT JOIN OLASItemDetailForOrdering ON Item.ID = OLASItemDetailForOrdering.NewItemCode
WHERE
  (
    (
      (Item.SalesDepartmentID)= 1
    )
    AND (
      (
        OLASItemDetailForOrdering.NewItemCode
      ) Is Null
    )
  );
