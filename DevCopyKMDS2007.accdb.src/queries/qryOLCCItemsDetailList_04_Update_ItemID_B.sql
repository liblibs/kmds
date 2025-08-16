UPDATE
  OLCCItemsDetailList
SET
  OLCCItemsDetailList.iTEMID = Left(
    [OLCCCODE],
    Len([OLCCCODE])-1
  )& "2"
WHERE
  (
    (
      (
        Right([olcccode], 1)
      )= "B"
    )
  );
