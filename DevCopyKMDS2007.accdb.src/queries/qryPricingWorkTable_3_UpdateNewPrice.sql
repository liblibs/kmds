UPDATE
  PricingWorkTable
SET
  PricingWorkTable.NewPrice = Round(
    (
      [Price] +(
        [Forms]![PricingTool]![txtPercent] * [PricingWorkTable].[Price]
      )* 0.01
    ),
    2
  );
