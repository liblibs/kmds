DELETE SalesPicturesTemp.FileName,
SalesPictures.*,
SalesPictures.ID
FROM
  SalesPicturesTemp
  RIGHT JOIN SalesPictures ON SalesPicturesTemp.FileName = SalesPictures.FileName
WHERE
  (
    (
      (SalesPicturesTemp.FileName) Is Null
    )
    AND (
      (SalesPictures.ID)> 0
    )
  );
