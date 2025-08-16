SELECT
  [vw_WebLogErrorCode-1].SetupDate,
  Trim([WebOrderID]) AS Invoice,
  Trim([WebOrderNotes]) AS OrderNotes,
  Trim([LicenseeName]) AS Customer,
  Trim([ErrorCode]) AS Error,
  Trim([Source]) AS Qty,
  OrdHead.InvoiceNo,
  OrdHead.InvQty,
  OrdHead.InvDollars,
  Trim([Subroutine]) AS eMail,
  Club.Contact,
  Club.Phone1
FROM
  (
    [vw_WebLogErrorCode-1]
    INNER JOIN OrdHead ON [vw_WebLogErrorCode-1].WebOrderID = OrdHead.InvoiceNo
  )
  INNER JOIN Club ON OrdHead.ClubID = Club.ID
ORDER BY
  [vw_WebLogErrorCode-1].SetupDate DESC;
