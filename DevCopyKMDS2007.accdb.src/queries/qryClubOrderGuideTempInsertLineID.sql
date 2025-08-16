PARAMETERS parClubID Long,
parLineID Long; INSERT INTO ClubOrderGuideTemp (ClubID, LineID)
SELECT
  [parClubID] AS Club,
  [parLineID] AS Line;
