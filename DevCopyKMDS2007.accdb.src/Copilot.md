# End of Day (EOD) Process Analysis

This document provides a detailed, step-by-step analysis of the `EOLog.Txt` file, correlating each log entry with the relevant VBA code, queries, and process logic in the MS Access End of Day (EOD) workflow.

---

## 1. Initialization and Setup

```
08:43:30 =========================================================
08:43:30 End of Day - Start
08:43:30 =========================================================
08:43:30 gstrMailDir = \\FS1\KMDS\
08:43:30 gstrNetworkDatabase = C:\KMDS\KMDS2007.ACCDR
08:43:30 Creating Folder : \\FS1\KMDS\Backup\2025\06\28\
```
- **Code:**  
  - `modEOD.bas`: `EndOFDayStream` logs the start, loads config if needed, and creates a backup folder for today.
  - Uses `FileSystemObject` to create year/month/day folders if missing.

---

## 2. SQL Backup

```
08:43:30 Dumping SQL Tables
08:43:30 MakeSQLBackup
```
- **Code:**  
  - `modDumpTable.bas`: `MakeSQLBackup` backs up SQL tables, likely using `DoCmd.TransferText` or similar.
  - Ensures data is preserved before changes.

---

## 3. Cleanup Old Files

```
08:43:36 Delete : C:\kmds\EOD File Transmit List.txt
08:43:36 Delete : K:LastNiteFiles.BAT
```
- **Code:**  
  - `modEOD.bas`: Uses `FileSystemObject.DeleteFile` to remove old transmit and batch files.

---

## 4. Main EOD Queries and Reports

All lines starting with `qryEOD...` are queries or reports executed by the EOD process.

- **Code:**  
  - `modEOD.bas`: `EOProcess` loops through all queries/reports with the EOD prefix, running each via `QueryExecute` or `ReportExecute`.
  - The number `-32764` indicates a report (see Access constants).

**Examples:**
```
08:43:36 qryEOD001InventoryValueBegin -32764
08:43:39 qryEOD010DeleteInvoHead5 5
08:43:39 qryEOD011DeleteInvoLine5 5
08:43:39 qryEOD021AppendCR1InvoLine 5
...
08:43:53 qryEOD200InventoryValueEnd -32764
```
- **Purpose:**  
  - Inventory snapshots, deleting/appending invoice data, updating costs, bottle deposits, cashier drops, generating sales/returns/productivity/club/restock reports, and final inventory value.
- **Queries:**  
  - Found in the `queries` folder, e.g., `qryEOD001InventoryValueBegin.bas`, `qryEOD010DeleteInvoHead5.bas`, etc.

---

## 5. Cleanup and Data Consolidation

```
08:43:54 qryEOD900ItemZeroSoldToday 5
08:43:55 qryEOD901SetSysfileCheckFlagTo0 5
08:43:55 qryEOD902DeleteItemsSold 5
08:43:55 qryEOD910CashierDropCountReset 5
08:43:55 qryEOD921PurgeInvoline_1_Mark_With_Z 5
08:44:00 qryEOD921PurgeInvoline_2_Delete_With_Z 5
08:44:01 qryEOD921PurgeKMDSLog500Days 5
08:44:01 qryEOD922SetItemSubCategoryIDTo0 5
```
- **Code:**  
  - More EOD queries for cleaning up zero-sold items, resetting flags, purging logs, resetting drop counts, and clearing subcategories.

---

## 6. Moving Register Files to Backup

```
08:44:01 Moving :\\FS1\KMDS\Mail\CR1.MDB to \\FS1\KMDS\Backup\2025\06\28\
...
08:44:01 Moving :\\FS1\KMDS\Mail\CR5.MDB to \\FS1\KMDS\Backup\2025\06\28\
```
- **Code:**  
  - `modEOD.bas`: Uses `FileSystemObject.MoveFile` to move each register’s data file to the backup folder, renaming with a timestamp.

---

## 7. Fintech and OLCC File Creation

```
08:44:01 Fintech = False
08:44:01 Create Daily Transmit files
08:44:02 Creating transmit file for inventory levels...
```
- **Code:**  
  - `modEOD.bas`: Checks if Fintech is enabled (`GetKMDSSettings("Fintech.Active", "False")`), runs `FinTechExport` if true.
  - Calls `OLCCCreateDailyTransmitFiles` and `OLCCInventFile` for OLCC compliance.

---

## 8. Lottery, Reindex, and Updates

```
08:44:02 Append Lottery Tickets Recon Info Begin
08:44:02 Append Lottery Tickets Recon Info End
08:44:02 Reindex Tables Begin
08:44:05 Reindex Tables End
08:44:05 Updating ItemMovement
08:44:07 Updating ItemCost
08:44:07 Updating Average Daily Sales and Assign DaysSupplyGroupIDs to Item Table...
```
- **Code:**  
  - `modEOD.bas`:  
    - `SQLPassThru "sp_LotteryScratch"` for lottery reconciliation.
    - `ReindexTables` for performance.
    - `SQLPassThru "sp_UpdateItemMovement"` and `"sp_UpdateItemCost"` for item movement/cost.
    - `AssignDaysSupplyGroupIDs` updates item supply groupings.

---

## 9. Customer Verification and File Transmission

```
08:44:14 ... VerifyCurrentKMDSCustomer: True
08:44:14 Sending OLCC Daily Transmit Files
08:44:14    Sending KMDS Daily Transmit Files
08:44:20 ...Complete Sending KMDS 
08:44:21 ...Complete Sending OLCC Daily Transmit Files
```
- **Code:**  
  - `modEOD.bas`:  
    - `VerifyCurrentKMDSCustomer` checks support status.
    - `OLCCDailyTransmit` and `WinSCPFTPTransmitKMDS` send files via FTP/SFTP.

---

## 10. Backup and Notification

```
08:44:21 MailSQLBackup
08:45:45 Backup Complete
08:45:45 SkipCopying c:\kmds\backup\KMDSBACK.zip to \\FS1\kmds\backup\KMDSSQL.ZIP
08:45:45 EODTextMessage
08:45:45 qryEOSalesForEMail
Error #94 Invalid use of Null : EOProcess
08:45:45 EODeMail Recap
08:45:45 EODeMailReports
08:45:45 Make Online Backup Batch File
08:45:45 Append UPCItems With ItemID
08:45:45 DropBoxInstalled = False
08:45:45 Getting OLCC ICE Item Files
08:45:45 .... complete!
08:45:45 =========================================================
08:45:45 End of Day = Finished Success
08:45:45 =========================================================
```
- **Code:**  
  - `modDumpTable.bas`: `MailSQLBackup` emails the SQL backup.
  - `modEOD.bas`:  
    - `EODTextMessage` sends a sales summary text/email.
    - `KMDSeMail` sends recap and reports.
    - `MakeOnlineBackupBatchFile` creates a batch file for online backup.
    - `AppendUPCItemsWithItemID` updates UPCs with ItemIDs.
    - Checks Dropbox status.
    - `GetOLCCItemFiles` downloads/imports OLCC ICE files.
    - Logs completion.

---

## 11. End of Month (EOM) Queries

```
10:08:05 qryEOM050DeleteNewpriceNulls 5
10:08:06 qryEOM099_Set_Cost_To_Price 5
10:08:07 qryEOM100ApplyNewPricesReport -32764
10:08:08 qryEOM101ApplyNewPricesUpdate 5
10:08:10 qryEOM110SalePricesReport -32764
10:08:10 qryEOM111SalePricesAppend 5
10:08:10 qryEOM201_Set_Price_To_CurPrice 5
10:08:10 qryEOM202MarkUpsAndDowns -32764
10:08:11 qryEOM203_Set_Cost_To_Price 5
10:08:12 qryEOM300WebPriceVariance -32764
10:08:14 qryEOM301WebScanVariance -32764
```
- **Code:**  
  - `modEOD.bas`:  
    - EOM queries/reports for price updates, sales, markups, and web variances.
    - Uses same `EOProcess`/`QueryExecute`/`ReportExecute` pattern.

---

## 12. Error Handling

```
Error #94 Invalid use of Null : EOProcess
```
- **Code:**  
  - `modEOD.bas`:  
    - Error in `EOProcess` (likely a missing or null value in a query result).
    - Error is logged and process continues.

---

## Summary Table

| Time      | Step/Query/Action                        | Description/Code Location                |
|-----------|------------------------------------------|------------------------------------------|
| 08:43:30  | Start, config, folder creation           | EOLog, LoadDefaults, FileSystemObject    |
| 08:43:30  | SQL backup                              | MakeSQLBackup                            |
| 08:43:36  | Delete old files                        | VBA Kill/DeleteFile                      |
| 08:43:36+ | EOD queries/reports                     | EOProcess, QueryExecute, ReportExecute   |
| 08:43:54+ | Cleanup, reset, purge                   | EOProcess, QueryExecute                  |
| 08:44:01  | Move register files to backup           | FileCopy/MoveFile                        |
| 08:44:01+ | Fintech/OLCC file creation              | FinTechExport, OLCCCreateDailyTransmitFiles, OLCCInventFile |
| 08:44:02+ | Lottery, reindex, update                | SQLPassThru, ReindexTables, AssignDaysSupplyGroupIDs |
| 08:44:14+ | Customer verify, file transmission      | VerifyCurrentKMDSCustomer, OLCCDailyTransmit, WinSCPFTPTransmitKMDS |
| 08:44:21+ | Backup, notifications                   | MailSQLBackup, EODTextMessage, KMDSeMail |
| 08:45:45  | Final steps, Dropbox, OLCC ICE files    | MakeOnlineBackupBatchFile, AppendUPCItemsWithItemID, GetOLCCItemFiles |
| 08:45:45  | End of Day finished                     | EOLog                                    |
| 10:08:05+ | End of Month queries/reports            | EOProcess, QueryExecute, ReportExecute   |

---

## Visual Flow (Simplified)

```plaintext
Start EOD
   |
   v
Load Config & Create Backup Folder
   |
   v
Make SQL Backup
   |
   v
Delete Old Files
   |
   v
Run EOD Queries/Reports (EOProcess)
   |
   v
Cleanup & Data Consolidation
   |
   v
Move Register Files to Backup
   |
   v
Create Transmit Files (Fintech/OLCC)
   |
   v
Lottery, Reindex, Update Inventory
   |
   v
Verify Customer & Transmit Files
   |
   v
Email Backup, Send Notifications
   |
   v
Final Steps (Backup, Dropbox, OLCC ICE)
   |
   v
End of Day Finished
```

---

## How to Use This Analysis

- **Each log line** corresponds to a function call, query, or report in your VBA code.
- **If you want to see the code for a specific step**, search for the function or query name in your modules or queries folder.
- **If you want to troubleshoot errors**, look for the error message and see which query or function was running at that time.

---

**If you want a deeper dive into any specific query, function, or error, or want a graphical flowchart, let me