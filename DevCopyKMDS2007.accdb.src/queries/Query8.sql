SELECT
  ReportForms.ReportName,
  ReportForms.ReportType
FROM
  ReportForms
WHERE
  (
    (
      (ReportForms.ReportType) Like "KMDS*"
    )
  );
