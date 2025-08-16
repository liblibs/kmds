UPDATE
  Item
SET
  Item.SubPack = upcecheckdigit([Subpack])
WHERE
  (
    (
      (
        Len([SubPack])
      )= 6
    )
  );
