SELECT
  Involine5.SetupDate AS [Date-Time],
  Involine5.InvoHeadID AS Invoice,
  Involine5.CashSysID AS Reigister,
  Involine5.Code AS ItemID,
  Involine5.InputCode AS UPCCode,
  Item.Description,
  Involine5.CashierID,
  Cashier.Name AS CashierName,
  Involine5.Quantity,
  Involine5.Price AS Amount
FROM
  (
    Involine5
    LEFT JOIN Cashier ON Involine5.CashierID = Cashier.ID
  )
  INNER JOIN Item ON Involine5.Code = Item.ID
WHERE
  (
    (
      (
        CStr([CODE])
      )= Involine5.InputCode
    )
    And (
      (Item.SalesDepartmentID)= 1
    )
    And (
      (Involine5.TransType)= "S"
    )
  );
