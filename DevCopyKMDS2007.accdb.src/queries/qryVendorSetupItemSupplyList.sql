SELECT
  qryVendorSetupItemSupplyListRAW.VendorID,
  Vendor.Name,
  Vendor.Contact,
  Vendor.Address1,
  Vendor.Address2,
  Vendor.City,
  Vendor.St,
  Vendor.Zip,
  Vendor.Phone1,
  Vendor.Phone2,
  Vendor.Phone3,
  Vendor.OrderCount,
  Vendor.PoAtCost,
  Vendor.PrintUpcOnPo,
  Vendor.POReportForm,
  Vendor.VendorNotes,
  qryVendorSetupItemSupplyListRAW.ID,
  qryVendorSetupItemSupplyListRAW.StockNumber,
  qryVendorSetupItemSupplyListRAW.Description,
  qryVendorSetupItemSupplyListRAW.SizeD,
  qryVendorSetupItemSupplyListRAW.Onhand,
  qryVendorSetupItemSupplyListRAW.Cost,
  qryVendorSetupItemSupplyListRAW.TodaysPrice,
  qryVendorSetupItemSupplyListRAW.OnReserve,
  qryVendorSetupItemSupplyListRAW.Avail,
  qryVendorSetupItemSupplyListRAW.OrderAt,
  qryVendorSetupItemSupplyListRAW.OrderTo,
  IIf(
    [Avail] <= [OrderAt], [OrderTo] - [Avail],
    0
  ) AS SOQ
FROM
  qryVendorSetupItemSupplyListRAW
  INNER JOIN Vendor ON qryVendorSetupItemSupplyListRAW.VendorID = Vendor.ID;
