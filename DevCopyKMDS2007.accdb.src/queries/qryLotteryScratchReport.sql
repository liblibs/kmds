SELECT
  LotteryScratchRecon.BusinessDate,
  LotteryScratchRecon.GameNumber,
  LotteryScratchRecon.ItemID,
  LotteryScratchRecon.BeginMeter,
  LotteryScratchRecon.EndMeter,
  LotteryScratchRecon.AddMeter,
  LotteryScratchRecon.AmountSold,
  LotteryScratchRecon.SetupDate,
  LotteryScratchRecon.ChangeDate,
  [BeginMeter] + [AddMeter] - [EndMeter] AS PhysicalAmount,
  [AmountSold] -(
    [BeginMeter] + [AddMeter] - [EndMeter]
  ) AS Variance,
  (
    [AmountSold] -(
      [BeginMeter] + [AddMeter] - [EndMeter]
    )
  )* [Price] AS VarianceDollars,
  Item.Description
FROM
  LotteryScratchRecon
  INNER JOIN Item ON LotteryScratchRecon.ItemID = Item.ID
WHERE
  (
    (
      (
        LotteryScratchRecon.BusinessDate
      )= [Forms]![LotteryScratch]![cmbBusinessDate]
    )
  );
