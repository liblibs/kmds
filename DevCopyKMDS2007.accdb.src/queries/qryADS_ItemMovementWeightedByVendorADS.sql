SELECT
  qryADS_ItemMovement.Code,
  [CurrMM] * IIf(
    IsNull([ADSCurrMM]),
    0,
    [ADSCurrMM]
  ) AS WADSCurrMM,
  [Jan] * IIf(
    IsNull([ADSJan]),
    0,
    [ADSJan]
  ) AS WADSJan,
  [Feb] * IIf(
    IsNull([ADSFeb]),
    0,
    [ADSFeb]
  ) AS WADSFeb,
  [Mar] * IIf(
    IsNull([ADSMar]),
    0,
    [ADSMar]
  ) AS WADSMar,
  [Apr] * IIf(
    IsNull([ADSApr]),
    0,
    [ADSApr]
  ) AS WADSApr,
  [May] * IIf(
    IsNull([ADSMay]),
    0,
    [ADSMay]
  ) AS WADSMay,
  [Jun] * IIf(
    IsNull([ADSJun]),
    0,
    [ADSJun]
  ) AS WADSJun,
  [Jul] * IIf(
    IsNull([ADSJul]),
    0,
    [ADSJul]
  ) AS WADSJul,
  [Aug] * IIf(
    IsNull([ADSAug]),
    0,
    [ADSAug]
  ) AS WADSAug,
  [Sep] * IIf(
    IsNull([ADSSep]),
    0,
    [ADSSep]
  ) AS WADSSep,
  [Oct] * IIf(
    IsNull([ADSOct]),
    0,
    [ADSOct]
  ) AS WADSOct,
  [Nov] * IIf(
    IsNull([ADSNov]),
    0,
    [ADSNov]
  ) AS WADSNov,
  [Dec] * IIf(
    IsNull([ADSDec]),
    0,
    [ADSDec]
  ) AS WADSDec
FROM
  qryADS_ItemMovement
  INNER JOIN (
    (
      Item
      INNER JOIN Vendor ON Item.VendorID = Vendor.ID
    )
    INNER JOIN AverageDaySupply ON Vendor.AverageDaySupplyID = AverageDaySupply.ID
  ) ON qryADS_ItemMovement.Code = Item.ID;
