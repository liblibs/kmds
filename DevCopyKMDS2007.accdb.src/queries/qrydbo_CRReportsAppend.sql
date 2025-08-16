INSERT INTO dbo_CRReports (ReportName, ZCopies)
SELECT
  CRReports.ReportName,
  CRReports.ZCopies
FROM
  CRReports
  LEFT JOIN dbo_CRReports ON CRReports.ReportName = dbo_CRReports.ReportName
WHERE
  (
    (
      (dbo_CRReports.ID) Is Null
    )
  );
