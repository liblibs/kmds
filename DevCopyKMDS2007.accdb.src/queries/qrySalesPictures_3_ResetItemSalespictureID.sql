UPDATE
  Item
  LEFT JOIN SalesPictures ON Item.PictureID = SalesPictures.ID
SET
  Item.PictureID = 0
WHERE
  (
    (
      (Item.PictureID)<> 0
    )
    AND (
      (SalesPictures.ID) Is Null
    )
  );
