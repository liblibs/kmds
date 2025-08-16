INSERT INTO SalesPictures (FileName)
SELECT
  SalesPicturesTemp.FileName
FROM
  SalesPictures
  RIGHT JOIN SalesPicturesTemp ON SalesPictures.FileName = SalesPicturesTemp.FileName
WHERE
  (
    (
      (SalesPictures.ID) Is Null
    )
  );
