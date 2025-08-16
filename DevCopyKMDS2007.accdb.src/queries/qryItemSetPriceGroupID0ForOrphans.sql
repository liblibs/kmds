UPDATE
  Item
  LEFT JOIN PriceGroup ON Item.PriceGroupID = PriceGroup.ID
SET
  Item.PriceGroupID = 0
WHERE
  (
    (
      (Item.PriceGroupID)<> 0
    )
    AND (
      (PriceGroup.ID) Is Null
    )
  );
