PARAMETERS parClubID Long,
parLineID Long;
UPDATE
  ClubOrderGuideTemp
SET
  ClubOrderGuideTemp.LineID = [LineID] + 1
WHERE
  (
    (
      (ClubOrderGuideTemp.LineID)>= [parLineID]
    )
    AND (
      (ClubOrderGuideTemp.ClubID)= [parClubID]
    )
  );
