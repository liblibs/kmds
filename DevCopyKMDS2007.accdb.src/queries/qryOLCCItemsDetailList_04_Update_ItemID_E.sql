UPDATE
  OLCCItemsDetailList
SET
  OLCCItemsDetailList.iTEMID = Left(
    [OLCCCODE],
    Len([OLCCCODE])-1
  )& "5"
WHERE
  (
    (
      (
        Right([olcccode], 1)
      )= "E"
    )
  );
