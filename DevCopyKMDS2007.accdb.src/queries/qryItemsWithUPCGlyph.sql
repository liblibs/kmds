SELECT
  qryUPCGlyph.ID AS UPCCode,
  Left([ItemD], 20) AS Descr,
  qryUPCGlyph.ItemID,
  ItemLabels.SalesDpt,
  ItemLabels.CurPrice,
  qryUPCGlyph.UPCG,
  ItemLabels.S,
  ItemLabels.Labels
FROM
  ItemLabels
  LEFT JOIN qryUPCGlyph ON ItemLabels.ID = qryUPCGlyph.ItemID
WHERE
  (
    (
      (ItemLabels.Labels)> 0
    )
  );
