SELECT
  Involine.SetupDate AS [Date-Time],
  Involine.InvoHeadID AS Invoice,
  Involine.CashSysID AS Reigister,
  Involine.Code AS ItemID,
  Involine.InputCode AS UPCCode,
  Involine.CashierID,
  Cashier.Name AS CashierName,
  Involine.Quantity,
  Involine.Price AS Amount
FROM
  (
    Involine
    LEFT JOIN Cashier ON Involine.CashierID = Cashier.ID
  )
  INNER JOIN Item ON Involine.Code = Item.ID
WHERE
  (
    (
      (Involine.SetupDate) Between Forms!PrintInventoryReport!txtDateRangeFrom
      And Forms!PrintInventoryReport!txtDateRangeTo
    )
    And (
      (
        CStr([CODE])
      )= Involine.InputCode
    )
    And (
      (Item.SalesDepartmentID)= 1
    )
    And (
      (Involine.TransType)= "S"
    )
  );
