INSERT INTO OLASiItemDetailForOrderingReportDate (F3, ReportDate)
SELECT
  ItemDetailForOrdering.F3,
  First(
    Right([F3], 10)
  ) AS ReportDate
FROM
  ItemDetailForOrdering
GROUP BY
  ItemDetailForOrdering.F3
HAVING
  (
    (
      (ItemDetailForOrdering.F3) Like "Report Date*"
    )
  );
