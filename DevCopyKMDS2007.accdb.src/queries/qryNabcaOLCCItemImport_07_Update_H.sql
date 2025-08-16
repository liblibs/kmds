UPDATE
  NabcaOlccItem
SET
  NabcaOlccItem.ItemID = CLng(
    Left([oldcode], 5)& "8"
  )
WHERE
  (
    (
      (NabcaOlccItem.OldCode) Like "*H"
    )
  );
