DELETE *
FROM
  dbo_KMDSLog
WHERE
  (
    (
      (dbo_KMDSLog.SetupDate)< Now()-500
    )
  );
