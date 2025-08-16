UPDATE
  (
    (
      ConvMIXMASTR
      INNER JOIN PriceGroup ON ConvMIXMASTR.Desc = PriceGroup.Description
    )
    INNER JOIN ConvITEMS ON ConvMIXMASTR.Code = ConvITEMS.MixmCode
  )
  INNER JOIN Item ON ConvITEMS.Code = Item.ID
SET
  Item.PriceGroupID = PriceGroup.ID;
