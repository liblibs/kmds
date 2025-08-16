UPDATE
  OLCCItemsDetailList
SET
  OLCCItemsDetailList.iTEMID = Left(
    [OLCCCODE],
    Len([OLCCCODE])-1
  )& "3"
WHERE
  (
    (
      (
        Right([olcccode], 1)
      )= "C"
    )
  );
