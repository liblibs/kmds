PARAMETERS parClubID Long,
parLineID Long; DELETE ClubOrderGuideTemp.*
FROM
  ClubOrderGuideTemp
WHERE
  (
    (
      (ClubOrderGuideTemp.ClubID)= [parClubID]
    )
    AND (
      (ClubOrderGuideTemp.LineID)= [parLineID]
    )
  );
