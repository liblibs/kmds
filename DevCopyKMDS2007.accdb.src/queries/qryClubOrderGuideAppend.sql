INSERT INTO ClubOrderGuides (
  ClubID, LineID, ItemID, Par, QuotedPrice,
  LiquorItem
)
SELECT
  ClubOrderGuideTemp.ClubID,
  ClubOrderGuideTemp.LineID,
  ClubOrderGuideTemp.ItemID,
  ClubOrderGuideTemp.Par,
  ClubOrderGuideTemp.QuotedPrice,
  IIf([SalesDepartmentID] = 1, 1, 0) AS LI
FROM
  ClubOrderGuideTemp
  INNER JOIN Item ON ClubOrderGuideTemp.ItemID = Item.ID
WHERE
  (
    (
      (ClubOrderGuideTemp.ClubID)= [parClubID]
    )
  );
