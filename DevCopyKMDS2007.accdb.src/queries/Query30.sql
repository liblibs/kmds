INSERT INTO Item (
  ID, Description, Price, SizeID, SalesDepartmentID,
  VendorID
)
SELECT
  GaribalidiInventory.ItemID AS ID,
  GaribalidiInventory.Description,
  GaribalidiInventory.Price,
  Right(
    [GaribalidiInventory].[itemid],
    1
  ) AS SizeID,
  1 AS SalesDepartmentID,
  1 AS VendorID
FROM
  GaribalidiInventory
  LEFT JOIN Item ON GaribalidiInventory.ItemID = Item.ID
WHERE
  (
    (
      (Item.ID) Is Null
    )
  );
