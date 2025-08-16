UPDATE
  NabcaOlccItem
SET
  NabcaOlccItem.ItemID = CLng(
    Left([oldcode], 5)& "6"
  )
WHERE
  (
    (
      (NabcaOlccItem.OldCode) Like "*F"
    )
  );
