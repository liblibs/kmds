INSERT INTO Item (
  ID, Description, SizeID, SalesDepartmentID,
  VendorID, Price, Cost
)
SELECT
  Sheet3.Code AS ID1,
  Sheet3.descr AS Description,
  Right([code], 1) AS SizeID,
  1 AS SalesDepartmentID,
  1 AS VendorID,
  Sheet3.Price,
  Sheet3.Price
FROM
  Item
  RIGHT JOIN Sheet3 ON Item.ID = Sheet3.Code
WHERE
  (
    (
      (Item.ID) Is Null
    )
  );
