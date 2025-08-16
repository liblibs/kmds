UPDATE
  Item
SET
  Item.SubPack = checkdigit11("0" & [subpack])
WHERE
  (
    (
      (
        Len([SubPack])
      )= 10
    )
  );
