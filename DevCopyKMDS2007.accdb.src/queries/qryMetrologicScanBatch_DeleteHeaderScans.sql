DELETE MetrologicScanbatch.ScanCode
FROM
  MetrologicScanbatch
WHERE
  (
    (
      (MetrologicScanbatch.ScanCode)= "-1"
    )
  );
