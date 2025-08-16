UPDATE
  NabcaOlccItem
SET
  NabcaOlccItem.ItemID = CLng(
    Left([oldcode], 5)& "2"
  )
WHERE
  (
    (
      (NabcaOlccItem.OldCode) Like "*B"
    )
  );
