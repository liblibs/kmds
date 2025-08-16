SELECT
  CashSys.ID,
  CashSys.ComputerName,
  Sysfile.Cr1Installed,
  Sysfile.Cr2Installed,
  Sysfile.Cr3Installed,
  Sysfile.Cr4Installed,
  Sysfile.Cr1Online,
  Sysfile.Cr2Online,
  Sysfile.Cr3Online,
  Sysfile.Cr4Online,
  Sysfile.Cr5Installed,
  Sysfile.Cr5Online
FROM
  Sysfile,
  CashSys
WHERE
  (
    (
      (CashSys.ID) Between 1
      And 5
    )
  );
