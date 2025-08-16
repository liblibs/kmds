PARAMETERS parClubID Long;
SELECT
  Club.ID,
  Club.OrderCount,
  Club.NonLiquor,
  Club.ClubNotes,
  Club.FintechAccount,
  Club.Licensee
FROM
  Club
WHERE
  (
    (
      (Club.ID)= [parClubID]
    )
  );
