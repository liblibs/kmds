SELECT
  qryVendorItemSupplyListRAW.VendorID,
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
  qryVendorItemSupplyListRAW.ID,
  qryVendorItemSupplyListRAW.StockNumber,
  qryVendorItemSupplyListRAW.Description,
  qryVendorItemSupplyListRAW.SizeD,
  qryVendorItemSupplyListRAW.Onhand,
  qryVendorItemSupplyListRAW.Cost,
  qryVendorItemSupplyListRAW.TodaysPrice,
  qryVendorItemSupplyListRAW.OnReserve,
  qryVendorItemSupplyListRAW.Avail,
  qryVendorItemSupplyListRAW.OrderAt,
  qryVendorItemSupplyListRAW.OrderTo,
  IIf(
    [Avail] <= [OrderAt], [OrderTo] - [Avail],
    0
  ) AS SOQ
FROM
  qryVendorItemSupplyListRAW
  INNER JOIN Vendor ON qryVendorItemSupplyListRAW.VendorID = Vendor.ID;
