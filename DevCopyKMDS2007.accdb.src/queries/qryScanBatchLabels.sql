SELECT
  ScanBatchQuantityLabels.UPCCode,
  UPCGlyph([UPCCode]) AS UPCG,
  ScanBatchQuantityLabels.Description,
  "*" & IIf(
    CLng([UPCCode])< 10,
    "0"
  )& CLng([UPCCode])& "*" AS F39,
  "* " & CLng([UPCCode])& "*" AS F39x
FROM
  ScanBatchQuantityLabels
ORDER BY
  ScanBatchQuantityLabels.UPCCode;
