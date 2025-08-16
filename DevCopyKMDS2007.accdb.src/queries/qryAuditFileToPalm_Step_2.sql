UPDATE
  InvoLineUPCMaxDate
  INNER JOIN UPCItem ON InvoLineUPCMaxDate.InputCode = UPCItem.ID
SET
  UPCItem.ChangeDate = InvoLineUPCMaxDate.maxofsetupdate;
