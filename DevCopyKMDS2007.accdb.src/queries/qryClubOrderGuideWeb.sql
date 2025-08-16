SELECT
  ClubOrderGuides.ClubID,
  ClubOrderGuides.OrderID,
  ClubOrderGuides.LineID,
  ClubOrderGuides.ItemID,
  ClubOrderGuides.Description,
  ClubOrderGuides.Par,
  ClubOrderGuides.QuotedPrice,
  Club.NonLiquor,
  ClubOrderGuides.LiquorItem,
  Club.Contact,
  Club.MaintainParLevels
FROM
  ClubOrderGuides
  INNER JOIN Club ON ClubOrderGuides.ClubID = Club.ID
WHERE
  (
    (
      (ClubOrderGuides.ClubID)= [Forms]![ClubOrderWebImport]![cmbLicensee]
    )
  );
