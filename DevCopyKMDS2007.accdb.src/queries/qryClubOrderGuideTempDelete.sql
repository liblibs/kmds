PARAMETERS parClubID Long; DELETE ClubOrderGuideTemp.ClubID,
ClubOrderGuideTemp.*
FROM
  ClubOrderGuideTemp
WHERE
  (
    (
      (ClubOrderGuideTemp.ClubID)= [parClubID]
    )
  );
