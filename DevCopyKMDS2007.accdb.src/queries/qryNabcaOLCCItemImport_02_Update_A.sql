UPDATE
  NabcaOlccItem
SET
  NabcaOlccItem.ItemID = CLng(
    Left([oldcode], 5)& "1"
  )
WHERE
  (
    (
      (NabcaOlccItem.OldCode) Like "*A"
    )
  );
