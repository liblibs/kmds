INSERT INTO UPCItem (ID, ItemID)
SELECT
  checkdigit11(
    Right([newUPC1], 11)
  ) AS ne,
  Query11.NewItemID AS Expr1
FROM
  Query11;
