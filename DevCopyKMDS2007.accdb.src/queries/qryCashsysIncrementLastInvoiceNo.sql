PARAMETERS parID Long;
UPDATE
  CashSys
SET
  CashSys.LastInvoiceNo = [LastInvoiceNo] + 1
WHERE
  (
    (
      (CashSys.ID)= [parID]
    )
  );
