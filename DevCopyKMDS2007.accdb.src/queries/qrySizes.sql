SELECT
  Size.ID,
  Size.Description,
  Size.Oz,
  Size.CaseQty,
  IIf(
    IsNull([qryItemsBySizeCounts].[count]),
    0,
    [qryItemsBySizeCounts].[count]
  ) AS Items
FROM
  qryItemsBySizeCounts
  LEFT JOIN [Size] ON qryItemsBySizeCounts.SizeID = Size.ID;
