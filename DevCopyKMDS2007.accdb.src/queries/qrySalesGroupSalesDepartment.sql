SELECT
  SalesGroup.ID,
  SalesGroup.Description,
  SalesDepartment.ID,
  SalesDepartment.Description,
  SalesDepartment.Discounts,
  SalesDepartment.NegativeOnhand,
  SalesDepartment.AllowDiscounts,
  SalesDepartment.IncludeInInventoryValueReports,
  SalesDepartment.TargetGP,
  SalesDepartment.AllowClubDiscounts,
  SalesDepartment.CatchAllItemID,
  SalesDepartment.SetupDate,
  SalesDepartment.ChangeDate
FROM
  SalesDepartment
  INNER JOIN SalesGroup ON SalesDepartment.SalesGroupID = SalesGroup.ID;
