SELECT
  ItemMovement.Description,
  ItemMovement.Code,
  [CurrMM] / Day(
    Now()
  ) AS ADSCurrMM,
  [Jan] / 31 AS ADSJan,
  [Feb] / 28 AS ADSFeb,
  [Mar] / 31 AS ADSMar,
  [Apr] / 30 AS ADSApr,
  [May] / 31 AS ADSMay,
  [Jun] / 30 AS ADSJun,
  [Jul] / 31 AS ADSJul,
  [Aug] / 31 AS ADSAug,
  [Sep] / 30 AS ADSSep,
  [Oct] / 31 AS ADSOct,
  [Nov] / 30 AS ADSNov,
  [Dec] / 31 AS ADSDec,
  Day(
    Now()
  ) AS today
FROM
  ItemMovement
WHERE
  (
    (
      (ItemMovement.Description)= "Sales - Total"
    )
  );
