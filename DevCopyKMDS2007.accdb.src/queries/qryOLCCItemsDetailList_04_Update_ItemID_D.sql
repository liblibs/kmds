UPDATE
  OLCCItemsDetailList
SET
  OLCCItemsDetailList.iTEMID = Left(
    [OLCCCODE],
    Len([OLCCCODE])-1
  )& "4"
WHERE
  (
    (
      (
        Right([olcccode], 1)
      )= "D"
    )
  );
