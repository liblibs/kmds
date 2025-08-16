SELECT
  qryUPCGlyph.UPCG,
  IIf(
    IsNull([ItemLabels].[UPCCodeMax])
    Or [ItemLabels].[ID] < 1000,
    [ItemLabels].[ID],
    [ItemLabels].[UPCCodeMax]
  ) AS BarCode,
  ItemLabels.ItemD AS Descr,
  ItemLabels.ID,
  ItemLabels.Labels,
  ItemLabels.SizeD,
  ItemLabels.CurPrice,
  ItemLabels.S,
  ItemLabels.SalesDpt,
  ItemLabels.Price,
  ItemLabels.VendorName,
  ItemLabels.SetupDate,
  ItemLabels.ChangeDate,
  IIf(
    [oz] > 0,
    Format([CurPrice] / [Oz], "#.00"" /oz"""),
    ""
  ) AS POZ
FROM
  (
    ItemLabels
    LEFT JOIN qryUPCGlyph ON ItemLabels.ID = qryUPCGlyph.ItemID
  )
  LEFT JOIN [Size] ON ItemLabels.SizeD = Size.Description
WHERE
  (
    (
      (ItemLabels.Labels)<> 0
    )
  );
