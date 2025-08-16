UPDATE
  OLCCItemsDetailList
SET
  OLCCItemsDetailList.iTEMID = Left(
    [OLCCCODE],
    Len([OLCCCODE])-1
  )& "6"
WHERE
  (
    (
      (
        Right([olcccode], 1)
      )= "F"
    )
  );
