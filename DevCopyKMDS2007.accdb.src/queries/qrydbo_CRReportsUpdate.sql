UPDATE
  CRReports
  INNER JOIN dbo_CRReports ON CRReports.ReportName = dbo_CRReports.ReportName
SET
  dbo_CRReports.ZCopies = CRReports.Zcopies;
