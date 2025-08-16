UPDATE
  OLCCItemsDetailList
SET
  OLCCItemsDetailList.iTEMID = Left(
    [OLCCCODE],
    Len([OLCCCODE])-1
  )& "8"
WHERE
  (
    (
      (
        Right([olcccode], 1)
      )= "H"
    )
  );
