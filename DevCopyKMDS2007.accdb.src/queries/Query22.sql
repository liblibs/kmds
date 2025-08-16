INSERT INTO Item (
  ID, SizeID, Description, Price, SalesDepartmentID,
  VendorID, Cost
)
SELECT
  Sheet2.[Item Code] AS ID,
  Right([Item Code], 1) AS SizeID,
  Sheet2.[Item Description] AS Description,
  Sheet2.[Current Retail Price] AS Price,
  1 AS SalesDepartmentID,
  1 AS VendorID,
  Sheet2.[Current Retail Price] AS Cost
FROM
  Sheet2
  LEFT JOIN Item ON Sheet2.[Item Code] = Item.ID
WHERE
  (
    (
      (Item.ID) Is Null
    )
  );
