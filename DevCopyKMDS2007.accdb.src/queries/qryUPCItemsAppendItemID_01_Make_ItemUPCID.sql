SELECT
  Item.id AS ItemID,
  CStr([ID]) AS UPCID INTO ItemUPCID
FROM
  Item;
