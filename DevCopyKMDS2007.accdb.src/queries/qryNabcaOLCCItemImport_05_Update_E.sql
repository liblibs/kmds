UPDATE
  NabcaOlccItem
SET
  NabcaOlccItem.ItemID = CLng(
    Left([oldcode], 5)& "5"
  )
WHERE
  (
    (
      (NabcaOlccItem.OldCode) Like "*E"
    )
  );
