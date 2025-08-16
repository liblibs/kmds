SELECT
  Club.ID AS ClubID,
  Club.Name AS ClubName
FROM
  Club
WHERE
  (
    (
      (Club.Name) Between [Forms]![PrintInventoryReport]![cboClubIDBegin]
      And [Forms]![PrintInventoryReport]![cboClubIDEnd]
    )
  );
