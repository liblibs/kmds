INSERT INTO UPCItem (ID, ItemID)
SELECT
  checkdigit11(
    Right([newUPC2], 11)
  ) AS ne,
  Query13.NewItemID AS Expr1
FROM
  Query13
WHERE
  (
    (
      ([Query13].[NewItemID])> 0
    )
  );
