INSERT INTO dbo_OLCCitem (
  Code, Description, Age, Proof, Price,
  UnitsPerCase, StockStatus, OldItemID,
  OldItemAlpha
)
SELECT
  Item_ICExpansion.Code,
  Item_ICExpansion.Description,
  Item_ICExpansion.Age,
  Item_ICExpansion.Proof,
  Item_ICExpansion.Price,
  Item_ICExpansion.UnitsPerCase,
  Item_ICExpansion.StockStatus,
  Item_ICExpansion.OldItemID,
  Item_ICExpansion.OldItemAlpha
FROM
  Item_ICExpansion
WHERE
  (
    (
      (Item_ICExpansion.Code) Is Not Null
    )
  );
