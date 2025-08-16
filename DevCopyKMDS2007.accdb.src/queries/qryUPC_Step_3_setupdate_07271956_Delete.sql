DELETE UPCItem.SetupDate
FROM
  UPCItem
WHERE
  (
    (
      (UPCItem.SetupDate)= #7/27/1956#
    )
  );
