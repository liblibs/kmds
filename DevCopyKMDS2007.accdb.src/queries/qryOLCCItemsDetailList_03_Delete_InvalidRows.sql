DELETE Len([OLCCCode]) AS Lid
FROM
  OLCCItemsDetailList
WHERE
  (
    (
      (
        Len([OLCCCode])
      )> 6
    )
  );
