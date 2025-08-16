INSERT INTO KMDSLinks (Caption, URL)
SELECT
  KMDSLinksLocal.Caption,
  KMDSLinksLocal.URL
FROM
  KMDSLinks
  RIGHT JOIN KMDSLinksLocal ON KMDSLinks.Caption = KMDSLinksLocal.Caption
WHERE
  (
    (
      (KMDSLinks.Caption) Is Null
    )
  );
