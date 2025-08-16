UPDATE
  NabcaOlccItem
SET
  NabcaOlccItem.ItemID = CLng(
    Left([oldcode], 5)& "4"
  )
WHERE
  (
    (
      (NabcaOlccItem.OldCode) Like "*D"
    )
  );
