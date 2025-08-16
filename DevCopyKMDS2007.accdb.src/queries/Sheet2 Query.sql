INSERT INTO Item (
  ID, Description, Price, SizeId, SalesDepartmentID,
  VendorID, Cost
)
SELECT
  Sheet2.[Code],
  Sheet2.[Description],
  Sheet2.[######] AS Price,
  Right([code], 1) AS SizeId,
  1 AS SalesDepartmentID,
  1 AS VendorID,
  Sheet2.[######] AS Cost
FROM
  Sheet2
  LEFT JOIN Item ON Sheet2.Code = Item.ID
WHERE
  (
    (
      (Item.ID) Is Null
    )
  );
