UPDATE
  Item
  INNER JOIN InvoLineMinSetupDate ON Item.ID = InvoLineMinSetupDate.Code
SET
  Item.SetupDate = [MinOfSetupDate]
WHERE
  (
    (
      (Item.SetupDate)> [MinOfSetupDate]
    )
  );
