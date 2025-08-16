UPDATE
  OrdHead,
  Sysfile
SET
  OrdHead.InvoiceNo = Sysfile.InvoiceCounter,
  OrdHead.InvDate = Now()
WHERE
  (
    (
      (OrdHead.ID)= [parOrdheadID]
    )
  );
