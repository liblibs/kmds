PARAMETERS parClubID Long; DELETE ClubOrderGuides.ClubID,
ClubOrderGuides.*
FROM
  ClubOrderGuides
WHERE
  (
    (
      (ClubOrderGuides.ClubID)= [parClubID]
    )
  );
