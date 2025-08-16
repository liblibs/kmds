UPDATE
  OLCCItemsDetailList
SET
  OLCCItemsDetailList.iTEMID = Left(
    [OLCCCODE],
    Len([OLCCCODE])-1
  )& "1"
WHERE
  (
    (
      (
        Right([olcccode], 1)
      )= "A"
    )
  );
