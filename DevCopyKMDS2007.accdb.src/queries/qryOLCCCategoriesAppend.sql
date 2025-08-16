INSERT INTO [OLCC Categories] (
  Category, Code, [Units Per Case],
  [size], Description
)
SELECT
  [OLCC Items from OLAS Website].Category,
  [OLCC Items from OLAS Website].ItemID,
  [OLCC Items from OLAS Website].UnitsPerCase,
  Right([itemid], 1) AS [size],
  [OLCC Items from OLAS Website].[Item Description]
FROM
  [OLCC Items from OLAS Website];
