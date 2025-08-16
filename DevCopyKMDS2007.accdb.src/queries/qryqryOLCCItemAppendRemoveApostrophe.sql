UPDATE
  Item_ICExpansion
SET
  Item_ICExpansion.Description = Replace([Description], "'", "")
WHERE
  (
    (
      (
        Replace([Description], "'", "")
      )<> [Description]
    )
  );
