PARAMETERS parClubID Long; INSERT INTO ClubOrderGuideTemp (
  LineId, ItemID, Par, ClubID, QuotedPrice,
  LiquorItem
)
SELECT
  ClubOrderGuides.LineID,
  ClubOrderGuides.ItemID,
  ClubOrderGuides.Par,
  ClubOrderGuides.ClubID,
  ClubOrderGuides.QuotedPrice,
  IIf([SalesDepartmentID] = 1, 1, 0) AS li
FROM
  ClubOrderGuides
  INNER JOIN Item ON ClubOrderGuides.ItemID = Item.ID
WHERE
  (
    (
      (ClubOrderGuides.ClubID)= [parClubID]
    )
  );
