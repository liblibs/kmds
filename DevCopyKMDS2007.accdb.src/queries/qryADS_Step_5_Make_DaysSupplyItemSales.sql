SELECT
  qryADS_ItemSold.Code,
  CLng(
    [sq] + IIf(
      IsNull([SumOfQuantity]),
      0,
      [SumOfQuantity] / [UnitsPerPack]
    )
  ) AS Sold INTO DaysSupplyItemSales
FROM
  qryADS_ItemSold
  LEFT JOIN qryADS_ItemSoldByPacks ON qryADS_ItemSold.Code = qryADS_ItemSoldByPacks.ID;
