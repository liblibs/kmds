Attribute VB_Name = "modKMDS"
'debug mode = False
'Production = True

Public Const gcfHandleErrors As Boolean = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'program docoumentation

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ver date       ' wha happened
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   2005-10-08 PurchaseOrdersOnFile            form_unload mailchanges
'               cmdDupe_Click                   remove resume
'               ItemDescription_BeforeUpdate    remove resume
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 150   2005-10-11  Report Club History by Bottle
'                       changed make table query uses qryClubInvoiceHistoryByInvoice
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 151   2005-10-11  PSC Support for club orders
'                   Audit Inventory Screen
'                   PSC Support for Audit Inventory Screen
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 152   2005-10-20  Lottery Ticket Recon ....
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 153   2005-10-20  Changed the vwItemsWithCurrentPricing
'                       by  to not force SQL Login Screen
'                       in lkupFormItemBySalesDepartmentID
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 154   2005-10-30  Changed the vwItemsWithCurrentPricing
'                       by  to not force SQL Login Screen
'                       in lkupFormItemByVendor and by Category and by Price Group
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 155   2005-10-21  added TablesAverageDailySupply and AverageDaySupplyLevels tables
'                   added Vendor.AverageDailySupplyID and zVendor.AverageDaySupplyLevels
'                   Maintain Average Day Supply Tables through the Vendor setup screen (C-2)
'                   Added columns to vwItemsWithCurrentPricing to drive the CreatePurchaseOrderWADS
'                   Create PO Screen - Removed Average Day Supply option, replaced With combo showing
'                       vendor's current WADS Table' if vendor has now WADS table assigned then
'                       Day's Supply Weighted Avg Daily Sales option is hidden
'
'                   Added Cashsys.pcChargePinPadComPort for users without pin pads (Saleem)
'
'160    2005-12-04
'                   added dynamic create of ftpord.scr and ftpinv.scr to push out the binary
'                   command when OLCC changed ftp servers...
'                   added ftp get fulprice.txt
'                   added import of fulprice.txt to compare prices of web
'                   added report qryEOM300WebPriceVariance to print compare at eom
'                   restored PO by ADS X Day's supply
'                   added reindex.sql at EOD
'                   changed ALL Connection strings to save password (SQL Login Box)
'                   "06 Inventory Count Sheets by Location" has "Order By" support
'                   D-3 InventoryReportParameters added report descriptions
'
'
'161    2005-12-08  added reindex routine at end of day...
'                   query runner has a new option to execute a stored proc
'                   lottery scratch not updating begin read on dates between 1st an 9th of month
'                   audit screen lets you preload with range of codes;sales dpt; locations etc...
'                   added Club Invoice Customer Copy to print by category...
'
'162    2005-12-11  version 161 didn't have any of the macros... (autoexec!) MapKDrive was not running
'
'163    2005-12-12  POS Pick List.. add 6 lines of "strMsg = strMsg & vbcrlf" to the end to push out
'                   vwItemsWithCurrentPricing - Primary Key on ID
'                   vwUPCItemsWithCurrentPricing Primary Key on UPCCode
'
'                   added   Clubs.Email
'                           Clubs.Licensee
'                           Agents.Email
'                           Agents.KMDSCustomer
'                   glnXFRAgentNumber to save XFR AgentNumber in transfer screen
'
'164    2005-12-28  added Audit report to print out changes from audit screen
'
'165    2006-01-07  removed 4 of the blue drop downs in the item screen because of the
'                       '8 concurrent queries' governer MSDE was choking on
'                   added option to only print variances in the audit screen
'                   added print preview to Label print because of formatting item barcodes
'
'166    2006-01-12  made 4 of the blue drop downs in the item screen not active until you double
'                   click them
'
'167    2006-01-16  Cashsys table added 4 fields.
'                       added support for 2 merchant id's liquor/related...
'                       added support to show avail during sale
'                       added support to play sounds during sale
'
'168    2006-01-16  added FetchJournals to Backup and View Journals
'                   added routine SetOrphanPriceGroupsToZero when deleting pricegroup
'                   fixed display of on order and on transfer in item screen
'
'170    2006-02-16  added RepriceClubOrder routine at formload of ClubOrderCreate
'                   added RepriceClubOrder routine at cmdPrint of ClubOrder
'                   because of 2 instances (eugene/beaumont) of extending at 0 dollars!

'172    2006-02-18  added reality check on the units per order..
'                   recalc units per order when orderTo < UnitsPerCase
'                        If OrderTo < UnitsPerOrder Then
'                            UnitsPerOrder = OrderTo - OrderAt
'                            UnitsPerOrder.BackColor = vbYellow
'                        End If
'                        If OrderTo > UnitsPerCase Then
'                            UnitsPerOrder = UnitsPerCase
'                            UnitsPerOrder.BackColor = vbYellow
'                        End If
'174    2006-03-02  added conditional over/short display in
'                   qryEOD040SalesByDepartment .. show over/short if drops or loans have
'                   been entered...
'                   added (restored) report qryEOD040SalesByTenderByRegister
'                   removed pesky msgbox in reportexecute letting user
'                   know a new report is being setup...
'
'175    2006-03-08  Store Transfers - added cost in Transfer Invoice
'                   and extended correctly when inputting from
'                   cash register (6 pack pricing was not right)..
'
'176    2006-03-14  added table UPCLibrary which is used by CR.EXE to default
'                   to description in related item setup...
'
'177    2006-04-17  added FTP InterStore Transfers
'                   Purchase Orders to include days/supply average daily sales
'                   Audit screen to support scanning
'
'178    2006-04-26  setup item from the transfer screen
'                   double click item in transfer screen to bring up item form
'                   added form in purchase order screen to set TPR's
'
'179    2006-05-06  mbPriceGroupRefresh added to prevent database crash when price group
'                   update occurred....
'                   transfer screen bug - selecting by description extends wrong columns
'
'                   added following price/cost change in transfers and purchase order lines
'                   in the Price_AfterUpdate event
'                    If MsgBox("Do you want to change the PRICE of this item?", vbYesNo) = vbYes Then
'                        MsgBox ("The PRICE must be changed in the Item Screen!")
'                        EditItem
'                    End If
'
'180    2006-05-26  added Open Accounts Reports
'                   added 42 Sales By Department by Category
'182    2006-06-??  added 43 Item Sales By Category
'183    2006-06-20  added on Other Purchase Orders to purchase order screen
'184    2006-07-07  removed the Create Purchase Order by Days Supply option
'185    2006-09-02   ((Item.SalesDepartmentID)<>1)); in qryUpdItemPriceGroupID
'186                  EOD-Sales by Register does not include register "Auth:" detail
'187    2006-09-15  qryPurchaseOrderLineUpdatePrice does not re-update cost! only price!
'
'188    2006-09-22  Club Order Notes default to the customer notes in CreateClubOrder routine
'                       !OrderNotes = strSaveNotes
'                   Club Order - Fixed "*" for special pricing...
'                   Purchase Orders - added FAX to display in Vendor Info Box
'       2006-09-24  changed from Now to datNow In post Transfer routine
'                   so that every row in Involine has same setupdate
'                   14 - business ledger has Transfer In / Out
'                   qry200 - liquor recap has Transfer In / Out
'
'       2006-10-11  New EOD Report... qryEOD053CancelledSalesByCashier
'                   Dispenser Setup... Export ALL Order guides to \\" & gstrCashsysInsLoc & "\KMDS\Clubs
'                   New reports...
'                       43 Item Sales By Category
'                       44 Store Transfers by Invoice
'                       45 Store Transfers by Invoice by Item
'                       46 Monthly Sales By Category
'
'                   New Form! SalesLeads - sub form off of G-5 system file.  maintain list displayed at register during a sale
'                   Dispenser reports - added sort by locations option...
'                   Changed Vendor Setup; added flag to turn off the alternating gray lines on PO Prints
'                   Changed Club Order Guide to not alternate gray lines on Par/Item ID
'                   Made CRReports Table local so I can send out new cash register reports in the Currver.zip
'                   Made ReportForms Table Local Access Table for above reason..
'                   Tightened up the PurchaseOrderWithPriceAndCost as per Lester @ BakerCity
'
'       2006-10-13  Added Previous Purchase orders to Vendor Screen -
'                   Added Print Previous PO's
'                   Added Undo a posted po routine to the vendor screen
'
'                   Added 2 columns to cash register system table... ShowRemoveItem / ShowDupeItem
'                       to allow control of new cash register functions... REMOVE / DUPE
'
'                   Added page to Club Set up screen to show previous invoices
'                   added report to reprint a club invoice from the club setup screen
'
'189    2006-11-20  added purchase order create by orderpoints with NO On Orders
'190
'191    2006-12-5   added interface CRView "Cash register remote viewer"
'                   'creates crview.ins in mail changes routine
'192    2006-12-6   added pcp.exe call from day end routine
'
'193    2007-01-20  added Daily Sales - Licensee...
'                       sysfile.OLCCFTPServer OLCCFTPUser OLCCFTPPassword OLCCFTPASCII
'                               OLCCFTPEndOfDaySend
'                       end of day routine makes call to OLCCCreateDailyTransmitFiles
'                                   and OLCCDailyTransmit
'                       added licensee table accessed through club set up screen
'                       implement KMDSFtp.exe - for more control during ftp sessions
'
'196    2007-02-16  added PalmDataDirectory Info to sysfile
'
'197    2007-03-10  added Palm query to reformat output/input file
'                   added sysfile.AutoDropCharge and AutoDropChecks
'                           sysfile form... end of day adds drops for checks/charges
'                           qryEOD030CashierAutodropCharges - qryEOD031CashierAutodropChecksReverse
'
'198    2007-03-24   changed club order screen at the save/print button
                    'for whatever reason.. this (me.refresh) was causing a crash..
                    'repair open data base error'
                     'Office2003 sp2 fixed it MOST of the time...
'       2007-03-31  added new form "ReAssignItemID" which uses a new stored procedure
'                   There is also a new itemid.exe used to 'mass' changes..
'                   The new form is launched the the d-7 audit screen...
'
'       2007-04-05  fixed autodropchecks/charge reverse queries; add 100 to cashier # instead of 1
'                   to prevent duplicate records in inv5line
'
'                   added form "ItemsByListBox" to replace menu option D-2; make item maintenance easier
'
'
'199    2007-04-07  Change end of day... wait until endo of stream to mail sql backup (MailSQLBackup)..
'                   run "kmdsdump" at start of routine (MakeSQLBackup)
'200    2007-04-16  added 51 report for today's sales
'                   fixed 'macro not found error' in club setup form in the form.dirty event
'
'201    2007-05-17  added Talon DVR Support; added change for Palm...
'
'202    2007-06-16  changed Item screen
'                       1. Price Groups when COST changes
'                       2. 6-Pack Scanning - don't allow QtyPack = 1 & QtyPackPrice = 0
'                       3. Add New Item has extended fields and duplicate displayed item...
'
'203    2007-06-23  Changed Item Movement sales total record from "Sales - Total" to "  SALES -- TOTAL" so
'                   it will be the first line in the item movement grid
'                   changed POS Pick list to print by Item ID (previously by line id)
'                   changed qryUPCItemMax999999 to only look at upcitems with a length of 12
'
'                   Changed qryCreateTransmitInventory to include 0 items
'
'204    2007-07-18  Added Cashsys.pcChargeUserNo to control what user number is sent to Active charge..
'                   restored menu option E7 Item Data sheet
'
'205    2007-08-02  Added shell to OLCCRetransmit.EXE from the A-5 screen...
'                   Changed Clubs form to not allow null or non numeric in licensee field
'                      (defaults to 0)
'                   added "Store Copy" to club invoices..
'                   Purchase Orders ... to facilitate scanning orders
'                       1. on new lines set tab stop to item id; on existing set to quantity shipped
'                       2. on new lines set quantity shipped to case quantity
'                   Audit Inventory ... added front/back/display/overflow areas for scanning
'                       1. tried to make doing inventory 'easier' by controlling execution flow when using
'                          counts by locations.
'                       2. added automatic variance report print of when posting
'                       3. Add New Button
'                       4. fixed NULL bug when adding new records by nz(Front) nz(back) nz(display) nz(overflow)
'                   added "53 Obsolete UPC Codes"
'
'                   added qryPurgeItems query to be run from the G-7 screen
'
'206    2007-10-8   cashsys pcChargePinPadTDES added pcCharge support for TDES pinpads
'
'207    2007-12-27  added support for pcCharge recon - find missing invoices
'                   added button in sysfile screen to refresh X-Reports List (dbo.crrreports)
'                   added update to Club Involine records.. sp_UpdateItemMovement step #14
'                               set salesgroupid to invoice #
'                   changed club screen to show previous invoice number
'                   added fields to club
'                           .StoreInvoices    - number of "Store Copy" Invoices
'                           .StoreInvoiceSort - sort by for above
'                           .OuncesColumn     - for customer invoices - variable NET price per /xx.xx ounces
'                           .ClubOrderGuideReport - 2-column - 1 column variation
'                   changed club setup/change screen for above fields
'                   changed store copy invoice report (and queries) for net price / x.xx ounce
'                   changed club invoice print dialog form to implement above changes
'
'                   changed purchase order screen... added 'Scan Mode' check box to stop at item id and order
'                           case quantity (as per darrel M0
'
'                   added report 54 club List
'                   added report 55 Vendor List
'                   added rsCashier.AllowShutDown... to be used for f2-lookup rights... Cashiers form and cr.exe
'                   added rsCashier.AllowDeleteDocument... allow delete club orders; po & transfers Cashier Forms
'                   added storing the transfer number in the involine.salesgroup field added to show history
'                   changed store transfer to display transfer # instead in inv.number
'                   added OLCC Claims screen - copied from Transfer out screen - allows entry of claim clode
'                   added Agents - previous transfers tab on display screen.. -
'                   changed PostTransfer routine to load salesgroup with onhand
'                   changed club order screen... added scan mode
'                   added setup/repair tab in sysfile form0.
'                   added google maps links to club and agents screens
'                   fixed eod060 - print the correct invoice number...
'                   update item purge queries...
'                       qryPurgeItems_Step_1_Delete_Involine_0_Qty_Not_R get rid of 0 qty items in involie
'                       qryPurgeItems_Step_2_Set_AVGTurns999 - don't purge order guide spacers
'                       qryPurgeItems_Step_3_Delete_AVGTurns999 do the deed
'
'209                fixed " Sales - Packs" for current month..
'                   added Date (mmm-yy) in order guide
'                   added vendor.POMethodDefault
'                   added sysfile screen... refresh tran types
'
'210  2008-01-09    changed qryClubHistoryByBottleByMonthMakeTable to include category
'                   changed qryClubInvoiceHistoryByInvoiceDetail refer to NonClubDiscntItemId for non -club disct $
'                   changed sp_updateMovement... added 15th step to delete out Itemmovement with 0
'
'
'211    2008-01-09  added purchase order report form show case/repack quantities
'                   addressed new months pricing issues by following changes
'                   changed Club Order Create to display "**** Reprint *****" when pricing has changed.
'                   added call to     SQLPassThru ("sp_Upd_TransferLinesWithCurrentPrice")
'                         in the 3 transfer screens
'                   added sp_UpdateItemCost - recalc cost for items in involine..
'                   pass Version Number in Backup to password protect and let me know version number
'
'213    2008-1-22   changed qryAuditFileFromPalm_Step_3 and qryAuditFileFromPalm_Step_4 to update
'                   the 4 location counts (front, back, display & overflow); use "ADD_INVOLINE_COST_AND_INV5LINE_COST" flag
'                   to determine file import spec (tab/comma) delimiter
'
'
'215    2008-2-11    changed query qrypcChargeDetailDebitAppend and qrypcChargeDetailTransAppend
'                       IIf(IsNull([Ticket]) Or [Ticket]="",0,[Ticket])
'
'                    changed PurchaseOrderCreate to call qryPurchaseOrderLineUpdateCost
'                       to fix liquor PO's cost created in prior month
'217    2008-04-11  added cashsys... pcChargeCommMethod,pcChargeIPAddress,pcChargePort,pcChargeEnableSSl
'                   pcChargeAllowHandEntere
'
'218    2008-07-02  removed all references to K:\ and will use UNC to stop sharing c:\ root
'                   integrate with windows authentication..
'                   1. club screen - licensee file
'                   2. club screen - excel dump
'                   3. eod = transmit inventory; sales
'                   4. setting network mail dir
'                   5. online backup
'                   6. olccftpget
'                   7. end of day
'                       a. looking for crx.mdb files
'                       b. move to archive directory
'                   8. reset connect string on involinex tables when server changes
'
'                   cash register program is now CR3.exe using KMDSIS.exe for credit cards...
'
'219    2008-08-16  separated cashier screen... Office sp3 caused screen to lock up because of
'                   too many subforms... added 'CashierRights' form for work around
'
'221    2008-09-15  added qryEOD029UpdateInvo5LineTransferToZero - cartons-pax set cost = 0
'                   also changed vw_Item_Cost_Recalc_Step_2 from >0 to <>0 to include negative items
'
'222    2008-11-16  fixed MailSQLBackup to from \\xxx\c\kmds\ to \\xxx\kmds
'                   added EODeMail to e-mail recap to owner
'
'223    2008-12-11  added Table KMDSLinks to allow websites - in new tab (URL) in G-5 screen
'                   added Table Settings to allow paired values - in new tab (Settings) in G-5 Screen
'
'224    2009-01-12  added support for PDF files.. acFormatSNP
'                           set gbSaveReportsAsPDF from sysfile.SaveReportsAsPDF
'
'225    2009=01-27  added function CashStatus to check if all active workstations are online
'
'233    2009-02-27  ADDED PROGRAM LOGGING
'
'234    2009-02-27  ADDED sysfile.ClubExportOGFormatInExcel
'
'251    2010-05-18  fixed non-numeric quantity in qryClubOrderXLAppend
'252    2010-06-22  relabled "Gross" to "Licensee" in qryEOD040 recap
'                   fixed calculation in "Last Night" in qryEOD200 recap
'253    2010-06-29  added macOnlineBackup AND create K:OnlineBackup.Bat with parameters for KMDSOLB.bat
                    'EOLog "Make Online Backup Batch File"
                    'MakeOnlineBackupBatchFile
'                   fixed calculation in "Last Night" in qryEOD200 recap
'300    2010-10-11  added GetOLCCItemFiles()
                    'OLCC provides the master files on the ftp website..
                    'this routine gets the files and imports them into SQL tables and loads stock status into
                    'item.stocknumber field... if stocknumber = "X" Or "S"pecial order then set vendor id =0
                    'added ItemSetup routine to prefill description/price stockstatus
                    'added button to purchase order screen to load form "OLCCItemsNotOnFile"
                    'added GetOLCCItemFiles to end of month routine

'301    2011-03-04  'fixed purchase order create by sales history.
                    'added days supply in purchase order lines grid ..

'302    2011-08-09  'added support for Kits
                    'need to add 2 records to CROBjects table.. -   Kit     Table
                    '                                               KitItem Table
'303    2011-09-17  ' updated Kit screen to add to item table
'                   ' made "CRObjects" a local table so "Kit" and "KitItem" did not need to be added to sql database
'
'
'304    2011-09-29  'GetOLCCItemFiles routine changed.
                    'step 7 Assign Items Vendor = 0 with stock number "X" or "S"
                    'REM Disable due to unpopular demand
                    'SQLPassThru "sp_OLCCItems_Step_4_Assign_Vendor_0"
'305    2011-11-12  ImportTextFile routine
'                   changed strFname = "\\" & ComputerName & "\C\KMDS\CRDetail\InvoLine.Txt"
'                   changed strFname = "\\" & ComputerName & "\KMDS\CRDetail\InvoLine.Txt"
'
'                   Over short is showing wrong in qryEOD040SalesByDepartment and qryEOD050SalesByAllCashiersByDepartment
'                   qryEOD030CashierAutodropChargesReverse - Changed inputcode from "Change" to Change "Autodrop"
'                   qryEOD031CashierAutodropChecksReverse  - Changed inputcode from "Change" to Change "Autodrop"
'
'                   added variable ftp Infor for transfers
'                        strFTPServer = GetKMDSSettings(KMDSFTPServer, "open ftp.kmds.us")
'                        strFTPUser = GetKMDSSettings(KMDSFTPServer, "Agents")
'                        strFTPPassword = GetKMDSSettings(KMDSFTPServer, "Ag3nt$")


'307    2011-12-15  Added F.5 Scan Batch utility to facilitate mobile inventory using a bluetooth scanner
'                    Added using an extra field in sysfile (ADD_INVOLINE_COST_AND_INV5LINE_COST) to control if
'                       Restock list includes club invoices sold

'31104  2013-01-31  F.5 confirmation inputbox when not on file scan occurs.
'31105  2013-01-31  Item setup bug fix... 2nd item set up had null UnitsPerCase
'                    'changed 2103-02-08
'                    rstemp!NumberOfCopies = 0
'                    rstemp.Update
'                    nNumberOfCopies = 0

'31210  2014-05-23  added Obsolete options to vendor, club and subcategories of items.
'31301  2014-10-15  added web orders..
'31302  2014-11-13  added PurchaseOrderLiquorWithOnShelf report...
'31403  2017-08-01  added support for Fintech payment processing; update brother printer driver...
'31407  2018-10-03  reformatted claims and purchase order report to widen id and g/p controls so as not
'                   to show #### marks in them.   also added 75 Inventory by Last Sold report ...
'31423              Form_PriceGroup changed delete button from macro to event to run the qryItemSetPriceGroupIDTo0 to prevent orphans



Option Compare Database
Option Explicit


Global gstrStoreName As String
Global gstrReportsDir As String
Global gstrMailDir As String
Global gstrNetworkDatabase As String
Global glsCashierName As String

Global glnOLCCVendor As Long
Global glnCashierID As Long
Global gdLoginTime As Date
Global gstrComputerName As String
Global glnCashsysID As Long

Global gstrIllegalFname
Global gstrPOSPrinterPort As String
Global gstrCashsysInsLoc As String

Global gstrWindowsUserName As String
Global gstrWindowsUserID As String
Global gstrWindowsUserPassword As String
Global gstrPassword As String


Global gstrAdjustReason As String
Global gstrLastArchivedReport As String
Global gstrFTPLogon As String
Global gstrFTPLogonKMDS As String
Global gstrPalmDataDirectoryInventoryTo As String
Global gstrPalmDataDirectoryInventoryFrom As String
Global gstrPalmDataDirectoryPurchaseTo As String
Global gstrPalmDataDirectoryPurchaseFrom As String
Global gstrFTPLogonWinSCPOLCC As String
Global gstrftpLogonWinSCPKMDS As String

Global gbAllowChangePrice As Boolean
Global gbAllowChangeOnhand As Boolean
Global gbAllowChangeCost As Boolean
Global gbAllowChangeOrderAt As Boolean
Global gbAllowChangeOrderTo As Boolean
Global gbAllowChangeDescription As Boolean
Global gbAllowHourlyWageView As Boolean
Global gbAllowDeleteDocument As Boolean
Global gbSaveReportsAsPDF As Boolean

Global gbUpdate As Boolean
Global gbEOProcessError As Boolean
Global gbRunMailChanges
Global gbTransferCartonsToPacks As Boolean
Global gbRentrakInstalled As Boolean
Global gbPrintRelatedOnOLCCInvoice As Boolean
Global gbOLCCFTPAscii As Boolean  'file transfer in ASCII or Binary?
Global gbOLCCFTPEndOfDaySend As Boolean
Global gbPalmCommaDelimited As Boolean
Global gbClubExportOGFormatInExcel As Boolean
Global gbIntegratedCreditCards As Boolean
Global gbKMDSFTPActive As Boolean
'added 31417
Global gbMaxLiquorDigitsIsSix As Boolean


Global glnColorInfo As Long
Global glnColorDropDown As Long
Global glnColorViewOnly As Long
Global glnColorTotals As Long
Global glnColorLockedForAll As Long

Global glnVersionNumber As Long
Global glnAgentNumber As Long
Global EODeMailFiles As New Collection
Global glnCashierIDForPasswordForm As Long
Global glngPurchaseOrderID As Long
Global glngScanBatchHeaderID As Long


'I wish i read this article years ago
'http://msdn.microsoft.com/en-us/library/ee358847(v=office.12).aspx#odc_ac2007_ta_ErrorHandlingAndDebuggingTipsForAccessVBAndVBA_Debugger


Dim mstrDB As String

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function Main()
'    MsgBox "Begin Main"
    LoadDefaults
    MapKDrive

    If gcfHandleErrors Then On Error Resume Next

    DoCmd.OpenForm "Login" ', acDesign, , , , acDialog
    If gcfHandleErrors Then On Error GoTo ErrLog1
    If CurrentProject.AllForms("Switchboard").IsLoaded Then
        If Forms("Switchboard").Visible = True Then
            Forms("Login").Visible = False
        End If
    End If
ExitSub:
    Exit Function
ErrLog:
    MsgBox Err.Number & " " & Err.Description, vbCritical, "Main"
ErrLog1:
Resume ExitSub
End Function

Public Sub ArchiveReport(ReportName As String, Optional ReportDescription As String)
Dim fso, msg, f
Dim fldr As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    If Not gcfHandleErrors Then On Error GoTo ErrLog
    If gstrReportsDir = "" Then LoadDefaults

    'create the directory for today e.g. "c:\KMDS\reports\2003\12\25"
    Set fso = CreateObject("Scripting.FileSystemObject")

    fldr = gstrReportsDir
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    fldr = fldr & Format(Now, "yyyy") & "\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    fldr = fldr & Format(Now, "mm") & "\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    fldr = fldr & Format(Now, "dd") & "\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If
    fldr = fldr & ReportName
    If Not IsNull(ReportDescription) Then
        fldr = fldr & " " & ReportDescription & " "
    End If
    fldr = fldr & Format(Now, "yyyy-mm-dd hh-nn-ss")

    If gbSaveReportsAsPDF Then
        fldr = fldr & ".PDF"
        DoCmd.OutputTo acOutputReport, ReportName, acFormatPDF, fldr
    Else
        fldr = fldr & ".snp"
        DoCmd.OutputTo acOutputReport, ReportName, acFormatSNP, fldr
    End If

    gstrLastArchivedReport = fldr
ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 2501 'no data to report
    Case Else
        MsgBox "Archive Report" & vbCrLf & Err.Number & " " & Err.Description
    End Select
    Resume ExitSub


End Sub

Public Function ArchiveReportShare(ReportName As String, Optional ReportDescription As String) As String
Dim fso, msg, f
Dim fldr As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    If gstrReportsDir = "" Then LoadDefaults

    'create the directory for today e.g. "c:\KMDS\reports\2003\12\25"
    Set fso = CreateObject("Scripting.FileSystemObject")

    fldr = gstrReportsDir
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    fldr = fldr & "EndOfDay" & "\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If


    fldr = fldr & Format(Now, "yyyy") & "\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    fldr = fldr & Format(Now, "mm") & "\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    fldr = fldr & Format(Now, "dd") & "\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If
    fldr = fldr & ReportName
    If Not IsNull(ReportDescription) Then
        fldr = fldr & " " & ReportDescription & " "
    End If

    fldr = fldr & Format(Now, "yyyy-mm-dd hh-nn-ss")
    On Error Resume Next
    If gbSaveReportsAsPDF Then
        fldr = fldr & ".pdf"
        ArchiveReportShare = fldr
        DoCmd.OutputTo acOutputReport, ReportName, acFormatPDF, fldr
    Else
        fldr = fldr & ".snp"
        ArchiveReportShare = fldr
        DoCmd.OutputTo acOutputReport, ReportName, acFormatSNP, fldr
    End If


ExitSub:

    Exit Function

ErrLog:
    Select Case Err.Number
    Case 2501 'no data to report
    Case Else
        MsgBox "ArchiveReportShare" & vbCrLf & Err.Number & " " & Err.Description
    End Select
    Resume ExitSub
End Function

Public Sub ArchiveReportDocument(ReportName As String, DocumentType As String, DocumentOwner As String, DocumentName As String)
Dim fso, msg, f
Dim fldr As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    If Not gcfHandleErrors Then On Error GoTo ErrLog
    If gstrReportsDir = "" Then LoadDefaults

    'create the directory for today e.g. "c:\KMDS\reports\DocumentType\DocumentOwner\DocumentName" & ".snp"
    '                                    "c:\KMDS\reports\ClubInvoice\The Taco House\11038.snp

    Set fso = CreateObject("Scripting.FileSystemObject")

    fldr = gstrReportsDir
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    fldr = fldr & DocumentType & "\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    fldr = fldr & DocumentOwner & "\"
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    If gbSaveReportsAsPDF Then
        fldr = fldr & DocumentName & ".pdf"
        DoCmd.OutputTo acOutputReport, ReportName, acFormatPDF, fldr
    Else
        fldr = fldr & DocumentName & ".snp"
        DoCmd.OutputTo acOutputReport, ReportName, acFormatSNP, fldr
    End If

ExitSub:
    Exit Sub

ErrLog:
    Select Case Err.Number
    Case 2501 'no data to report
    Case Else
        MsgBox "ArchiveReportDocument" & vbCrLf & Err.Number & " " & Err.Description
    End Select
    Resume ExitSub
End Sub


Public Function CashStatus() As String
    'CHECK THE STATUS OF EACH WORKSTATION BY LOOKING FOR A SHARE CALLED 'KMDS'


    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")

    Dim fso As New FileSystemObject
    Dim bAllOnline As Boolean
    Dim strMsg As String
    Dim strCashRegisters As String

    If gcfHandleErrors Then On Error GoTo ErrLog

    bAllOnline = True
    Screen.MousePointer = 11 'busy

    rstemp.Open "qryCashsysOnLine", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
    With rstemp
        Do While Not .EOF
            strCashRegisters = strCashRegisters & !ComputerName & vbCrLf
            If fso.FolderExists("\\" & !ComputerName & "\kmds") Then

            Else

                bAllOnline = False
                strMsg = strMsg & !ComputerName & "    OFFLINE !!!" & vbCrLf
                Exit Do
            End If
            .MoveNext
        Loop
    End With

ExitSub:
    rstemp.Close
    Set rstemp = Nothing
    Set fso = Nothing
    If Not bAllOnline Then
        strMsg = "NOT OK..Problems with checking ACTIVE computers" & vbCrLf & strMsg
    Else
        strMsg = strCashRegisters
    End If
    CashStatus = strMsg
    Screen.MousePointer = 1 'normal
Exit Function

ErrLog:
    strMsg = strMsg & "Error #" & Err.Number & " " & Err.Description & vbCrLf
    bAllOnline = False
    Resume ExitSub
End Function


Public Function CopyUpdates()

    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")

    Dim fso As New FileSystemObject

    Dim bAllOnline As Boolean
    Dim strMsg As String
    Dim strSource As String
    Dim strTarget As String
    Dim varReturn As Variant

    If gcfHandleErrors Then On Error GoTo ErrLog

    If gstrCashsysInsLoc = "" Then LoadDefaults

    strMsg = CashStatus()


    If InStr(strMsg, "Problems") Then
        MsgBox strMsg
        Exit Function
    End If

    strSource = "\\" & gstrCashsysInsLoc & "\kmds\versions\*.*"

    Screen.MousePointer = 11 'busy
    rstemp.Open "qryCashsysOnLine", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
    Do While Not rstemp.EOF
        If gstrCashsysInsLoc <> rstemp!ComputerName Then
            varReturn = SysCmd(acSysCmdSetStatus, "Copying " & strSource & " to " & strTarget)
            strTarget = "\\" & rstemp!ComputerName & "\kmds\versions\"
            fso.CopyFile strSource, strTarget, True
        End If
        rstemp.MoveNext
    Loop

ExitSub:
    rstemp.Close
    Set rstemp = Nothing
    Set fso = Nothing
    Screen.MousePointer = 1 'normal
    MsgBox strMsg
    varReturn = SysCmd(acSysCmdClearStatus)

Exit Function

ErrLog:
    strMsg = strMsg & " Error #" & Err.Number & " " & Err.Description & vbCrLf
    strMsg = strMsg & "Source: " & strSource & vbCrLf
    strMsg = strMsg & "Target: " & strTarget & vbCrLf

    bAllOnline = False
    Resume ExitSub

End Function


Public Function CheckDirty(MyForm As Form)
    If MyForm.Form.Dirty Then
        MyForm.Form.Dirty = False
        MyForm.Refresh
    End If
End Function


Public Sub EMsgBox(MyForm As Form, MySub As String)

Dim strMsg As String
    If gcfHandleErrors Then On Error Resume Next
    strMsg = Err.Number & " " & Err.Description & vbCrLf & vbCrLf
    strMsg = strMsg & "Form: " & MyForm.Name & vbCrLf
    strMsg = strMsg & " Sub: " & MySub
    'If Err.Number > 0 Then
        MsgBox strMsg, vbInformation, "KMDS Error"
    'End If

End Sub

Public Sub fMsgBox(MyFunction As String)

Dim strMsg As String
    If gcfHandleErrors Then On Error Resume Next
    strMsg = Err.Number & " " & Err.Description & vbCrLf & vbCrLf
    strMsg = strMsg & " Function: " & MyFunction
    'If Err.Number > 0 Then
        MsgBox strMsg, vbInformation, "KMDS Error"
    'End If

End Sub


Public Function GetStoreName() As String
    If gstrStoreName = "" Then LoadDefaults
    GetStoreName = gstrStoreName
End Function





Public Function KMDSLogin(nID As Long, strPassword As String) As Boolean

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object

Dim ndays As Integer
glnCashierID = 0
glsCashierName = ""
gstrPassword = strPassword

If gcfHandleErrors Then On Error GoTo ErrLog
    KMDSLogin = False
    'ArchiveReport "PurchaseOrderCreate"
    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryCashierLogin"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parID", adInteger, adParamInput)
        .Parameters("parID") = nID

        .Parameters.Append .CreateParameter("parPassword", adBSTR, adParamInput)
        .Parameters("parPassword") = KMDSEncrypt(strPassword)
        '.Parameters("parPassword") = (strPassword)
    End With

    rstemp.Open comQDF, , adOpenStatic, adLockOptimistic ' , adLockReadOnly
        If rstemp.BOF And rstemp.EOF Then
            Set comQDF = Nothing
            Set comQDF = CreateObject("ADODB.COMMAND")
            With comQDF
                rstemp.Close
                .ActiveConnection = CurrentProject.Connection
                .CommandText = "qryCashierLogin"
                .CommandType = adCmdStoredProc
                'ADO Numeric Datatypes are very particular
                .Parameters.Append .CreateParameter("parID", adInteger, adParamInput)
                .Parameters("parID") = nID

                .Parameters.Append .CreateParameter("parPassword", adBSTR, adParamInput)
                '.Parameters("parPassword") = KMDSEncrypt(strPassword)
                .Parameters("parPassword") = (strPassword)
                rstemp.Open comQDF, , adOpenStatic, adLockReadOnly
            End With
        End If
    With rstemp
        If .BOF And .EOF Then
        Else
            KMDSLogin = True
            glnCashierID = !ID

            glsCashierName = !Name
            'MsgBox "Welcome!", , glnCashierID & " " & glsCashierName
            gdLoginTime = Now
            gbAllowChangePrice = !AllowChangePrice
            gbAllowChangeOnhand = !AllowChangeOnhand
            gbAllowChangeCost = !AllowChangeCost
            gbAllowChangeOrderAt = !AllowChangeOrderAt
            gbAllowChangeOrderTo = !AllowChangeOrderTo
            gbAllowChangeDescription = !AllowChangeDescription
            gbAllowHourlyWageView = !AllowHourlyWageView
            gbAllowDeleteDocument = !AllowDeleteDocument
            If gbIntegratedCreditCards Then
                If !AccountLocked Then
                    If DateDiff("n", !AccountLockedDate, Now) < 4 Then
                        MsgBox "Account locked; please wait for " & 4 - (DateDiff("n", !AccountLockedDate, Now)) & " minutes"
                        KMDSLogin = False
                        GoTo ExitFunc
                    Else
                        SQLPassThru ("sp_CashierAccountLockedReset " & CStr(nID))
                    End If
                End If
                If !PasswordDate + 80 < Now Then
                    If !PasswordDate + 90 < Now Then
                        MsgBox "Can't log in because your password has expired!"
                        MsgBox "Password has expired... " & vbCrLf & "You are required to change your password!!!"
                        glnCashierIDForPasswordForm = glnCashierID
                        DoCmd.OpenForm "CashierPassword", acNormal, , , , acDialog
                        KMDSLogin = False
                    Else
                        ndays = (!PasswordDate + 90) - Now
                        If ndays >= 1 Then
                            MsgBox "Your password will expire in " & ndays & " days"
                        Else
                            MsgBox "Your password expires TODAY!" & vbCrLf & "You are required to change your password by tomorrow"
                        End If
                    End If
                End If
            End If
        End If
        .Close
    End With

ExitFunc:
    Set comQDF = Nothing
    Set rstemp = Nothing
    Exit Function
ErrLog:

    MsgBox "Error"
    Resume ExitFunc

End Function

Public Sub LoadDefaults()
'    MsgBox "begin Loaddefaults"
    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")


    Dim strSQL As String

    If gcfHandleErrors Then On Error GoTo ErrLog
    rstemp.Open "SysFile", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    With rstemp

        If IsNull(!ReportsDir) Or !ReportsDir = "" Then
            MsgBox "Setting the reports directory to C:\KMDS\Reports\"
            !ReportsDir = "C:\KMDS\Reports\"
            .Update

        End If
        If IsNull(!NetworkMailDir) Or !NetworkMailDir = "" Or !NetworkMailDir = "K:\KMDS\" Then
            MsgBox "Setting the NetworkMailDir directory to \\" & gstrCashsysInsLoc & "\KMDS\"
            !NetworkMailDir = "\\" & gstrCashsysInsLoc & "\KMDS\"
            .Update

        End If

        If IsNull(!NetworkDatabase) Or !NetworkDatabase = "" Then
            MsgBox "Setting the NetworkDatabase directory to C:\KMDS\KMDSSQL.MDE"
            !NetworkDatabase = "C:\KMDS\KMDSSQL.MDE"
            .Update
        End If

        gstrStoreName = !AgentNumber & " " & !Name
        gstrReportsDir = !ReportsDir
        gstrMailDir = !NetworkMailDir
        gstrNetworkDatabase = !NetworkDatabase
        glnOLCCVendor = !OLCCVendor

        If IsNull(!ColorInfo) Then
            MsgBox "Setting Default colors!"
            !ColorInfo = 8454143
            !ColorDropDown = 16777088
            !ColorViewOnly = 8421631
            !ColorTotals = 8453888
            !ColorLockedForAll = 255
            .Update
        End If

        glnColorInfo = !ColorInfo
        glnColorDropDown = !ColorDropDown
        glnColorViewOnly = !ColorViewOnly
        glnColorTotals = !ColorTotals
        glnColorLockedForAll = !ColorLockedForAll

        If IsNull(!CashsysInsLoc) Then
            MsgBox "Setting CashsysInsLoc (SQL Server Workstation)!"
            !CashsysInsLoc = "FS1"
            .Update
        End If
        gstrCashsysInsLoc = !CashsysInsLoc
        glnAgentNumber = !AgentNumber
        gbTransferCartonsToPacks = !TransferCartonsToPacks
        gbRentrakInstalled = !RentrakInstalled
        gbPrintRelatedOnOLCCInvoice = !PrintRelatedOnOLCCInvoice
        '2007-01-20
        gstrFTPLogon = !OLCCFTPServer & vbCrLf
        gstrFTPLogon = gstrFTPLogon & !OLCCFTPUser & vbCrLf
        gstrFTPLogon = gstrFTPLogon & !OLCCFTPPassword & vbCrLf


        gbOLCCFTPAscii = !OLCCFTPASCII
        gbOLCCFTPEndOfDaySend = !OLCCFTPEndOfDaySend

        'If not gcfHandleErrors then Debug.Print gstrFTPLogon; " "; gbOLCCFTPAscii
        gstrPalmDataDirectoryInventoryTo = Nz(!PalmDataDirectoryInventoryTo, "")
        gstrPalmDataDirectoryInventoryFrom = Nz(!PalmDataDirectoryInventoryFrom, "")
        gstrPalmDataDirectoryPurchaseTo = Nz(!PalmDataDirectoryPurchaseTo, "")
        gstrPalmDataDirectoryPurchaseFrom = Nz(!PalmDataDirectoryPurchaseFrom, "")
        gbPalmCommaDelimited = !ADD_INVOLINE_COST_AND_INV5LINE_COST
        gbSaveReportsAsPDF = !SaveReportsAsPDF
        gbClubExportOGFormatInExcel = !ClubExportOGFormatInExcel

        'open ftp://Agent1181:Agent1181@ftp.olcc.state.or.us/ -rawsettings ProxyPort=0
        gstrFTPLogonWinSCPOLCC = "open ftp://User:Password@ftpsite/ -rawsettings ProxyPort=0"
        gstrFTPLogonWinSCPOLCC = Replace(gstrFTPLogonWinSCPOLCC, "User", !OLCCFTPUser)
        gstrFTPLogonWinSCPOLCC = Replace(gstrFTPLogonWinSCPOLCC, "Password", !OLCCFTPPassword)
        gstrFTPLogonWinSCPOLCC = Replace(gstrFTPLogonWinSCPOLCC, "ftpsite", !OLCCFTPServer)

        'open ftp://Agents:Ag3nt%24@ftp.kmds.us/ -rawsettings ProxyPort=0
        gstrftpLogonWinSCPKMDS = "open ftp://User:Password@ftpsite/ -rawsettings ProxyPort=0"
        gstrftpLogonWinSCPKMDS = Replace(gstrftpLogonWinSCPKMDS, "User", GetKMDSSettings("KMDSFTPUser", "Agents"))
        gstrftpLogonWinSCPKMDS = Replace(gstrftpLogonWinSCPKMDS, "Password", GetKMDSSettings("KMDSFTPPassword", "Ag3nt$"))
        gstrftpLogonWinSCPKMDS = Replace(gstrftpLogonWinSCPKMDS, "ftpsite", GetKMDSSettings("KMDSFTPServer", "ftp.kmds.us"))
        gstrftpLogonWinSCPKMDS = Replace(gstrftpLogonWinSCPKMDS, "$", "%24") ' replace special characters with ASCII equivalent



        .Close
    End With
    Set rstemp = Nothing

    Set rstemp = CreateObject("ADODB.Recordset")
    'MsgBox "Open Cashsys"

    gstrComputerName = MyComputerName
    strSQL = "SELECT CashSys.ComputerName,PrinterPort, ID, pcChargeInstalled  "
    strSQL = strSQL & "FROM CashSys "
    strSQL = strSQL & "WHERE (((CashSys.ComputerName) = '" & gstrComputerName & "'))"
    rstemp.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rstemp.BOF And rstemp.EOF Then
        MsgBox "There is no Cash Register System record for this machine: " & gstrComputerName
    Else
        glnCashsysID = rstemp!ID
        gstrPOSPrinterPort = Nz(rstemp!PrinterPort, "No Printer Port for this cash register ")
    End If
    rstemp.Close
    Set rstemp = Nothing

    Set rstemp = CreateObject("ADODB.Recordset")
    strSQL = "SELECT CashSys.ComputerName,PrinterPort, ID, pcChargeInstalled  "
    strSQL = strSQL & "FROM CashSys "
    strSQL = strSQL & "WHERE ((CashSys.pcChargeInstalled) <>0)"
    rstemp.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rstemp.BOF And rstemp.EOF Then
        gbIntegratedCreditCards = False
    Else
        gbIntegratedCreditCards = True
    End If
    rstemp.Close
    Set rstemp = Nothing





    'since club/vendor names are used in the file structure
    'they can not have illegal file names!

    gstrIllegalFname = "\/:*?<>|" & Chr(34)
    'If not gcfHandleErrors then Debug.Print gstrIllegalFname

    'gcfHandleErrors = GetKMDSSettings("DebugMode", "False")
    gbKMDSFTPActive = GetKMDSSettings("KMDSFTPActive", "False")

    gstrFTPLogonKMDS = GetKMDSSettings("KMDSFTPServer", "ftp.kmds.us") & vbCrLf
    gstrFTPLogonKMDS = gstrFTPLogonKMDS & GetKMDSSettings("KMDSFTPUser", "Agents") & vbCrLf
    gstrFTPLogonKMDS = gstrFTPLogonKMDS & GetKMDSSettings("KMDSFTPPassword", "Ag3nt$") & vbCrLf

    'added 31417
    gbMaxLiquorDigitsIsSix = GetKMDSSettings("MaxLiquorDigitsIsSix", "False") 'false is 1000 - to 99999

    Set rstemp = CreateObject("ADODB.Recordset")
    strSQL = "SELECT * "
    strSQL = strSQL & "FROM ScanBatchHeader "
    With rstemp
        .Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If .BOF And .EOF Then
            'add a header recordd//
            .AddNew
            !CashierID = 99
            .Update
            glngScanBatchHeaderID = !ID

        Else
            .MoveFirst
            glngScanBatchHeaderID = !ID

        End If
        Debug.Print glngScanBatchHeaderID

        .Close
    End With
    Set rstemp = Nothing

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "LoadDefaults"

    Resume ExitSub
End Sub

Public Function FormLogin(nCashierID As Long, MyForm As Form) As Boolean

    'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command

Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim strSQLString As String

If gcfHandleErrors Then On Error GoTo ErrLog
    FormLogin = True
    'ArchiveReport "PurchaseOrderCreate"
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        '2011-12-11
        .CommandText = "qryCashiersAndProgramsFormLogin"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parCashierID", adInteger, adParamInput)
        .Parameters("parCashierID") = nCashierID
        .Parameters.Append .CreateParameter("parProgramName", adLongVarChar, adParamInput, Len(MyForm.Name) + 1)
        .Parameters("parProgramName") = MyForm.Name
    End With

    rstemp.Open comQDF, , adOpenStatic, adLockReadOnly
    With rstemp
        If .RecordCount = 0 Then
            MsgBox "Cashier #" & nCashierID & " does not have permission to run this program"
        Else
            FormLogin = False
            SetFormSize MyForm
            strSQLString = "sp_KMDSLog "
            strSQLString = strSQLString & MyForm.Name & ", "    '@ProgramName
            strSQLString = strSQLString & "I, " '@InOut
            strSQLString = strSQLString & nCashierID & ", " '@CashierID
            strSQLString = strSQLString & glnCashsysID & ", " '@CashsysId
            strSQLString = strSQLString & "0, " '@ErrNumber
            strSQLString = strSQLString & "OK " '@ErrDescription
            SQLPassThru strSQLString
        End If
        .Close
    End With

ExitFunc:
    Set comQDF = Nothing
    Set rstemp = Nothing
    Exit Function
ErrLog:
    MsgBox "Error" & Err.Number & " " & Err.Description

    Resume ExitFunc

End Function

Public Function FormLogout(nCashierID As Long, MyForm As Form) As Boolean
Dim strSQLString As String

If gcfHandleErrors Then On Error GoTo ErrLog
    strSQLString = "sp_KMDSLog "
    strSQLString = strSQLString & MyForm.Name & ", "    '@ProgramName
    strSQLString = strSQLString & "O, " '@InOut
    strSQLString = strSQLString & nCashierID & ", " '@CashierID
    strSQLString = strSQLString & glnCashsysID & ", " '@CashsysId
    strSQLString = strSQLString & "0, " '@ErrNumber
    strSQLString = strSQLString & "OK " '@ErrDescription
    SQLPassThru strSQLString

ExitFunc:
    Exit Function
ErrLog:
    MsgBox "Error" & Err.Number & " " & Err.Description
    Resume ExitFunc
End Function

Public Sub SetFormSize(MyForm As Form)
    If gcfHandleErrors Then On Error Resume Next
    If Not gcfHandleErrors Then On Error Resume Next
    MyForm.InsideHeight = MyForm.FormHeader.Height + MyForm.Detail.Height + MyForm.FormFooter.Height
    MyForm.InsideWidth = MyForm.Form.Width + 325
    MyForm.Caption = MyForm.Caption & " KMDS Software © "
End Sub

Public Sub AssignClubInvoiceNO(OrdheadID As Long)
    SQLPassThru "sp_Assign_InvoiceNO " & OrdheadID
End Sub


Public Sub MapKDrive()
    'make sure "K" is mapped

    If gstrCashsysInsLoc = gstrComputerName Then
        WaitForExecCmd (CStr("c:\kmds\KMDSNetSub.bat " & gstrCashsysInsLoc))
    Else
        WaitForExecCmd (CStr("c:\kmds\KMDSNet.bat " & gstrCashsysInsLoc))
    End If
End Sub

Public Function MyComputerName() As String
    Dim varComp As String * 20
    Dim lngReturn As Long
    Dim lngN As Long
    Dim nI As Integer

    'get the computer name... this is the cash register name
    lngN = 20
    lngReturn = GetComputerName(varComp, lngN)
    nI = 1
    Do While Asc(Mid(varComp, nI, 1)) > 0
        MyComputerName = MyComputerName & Mid(varComp, nI, 1)
        nI = nI + 1
    Loop

End Function

Public Sub PriceGroupUpdate(PriceGroupID As Long, NewPrice As Currency, QtyPack As Integer, QtyPackPrice As Currency, Cost As Currency)

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

        With comQDF
            .ActiveConnection = CurrentProject.Connection
            .CommandText = "qryUpdItemPriceGroupID"
            .CommandType = adCmdStoredProc
            'ADO Numeric Datatypes are very particular
            .Parameters.Append .CreateParameter("parPriceGroupID", adInteger, adParamInput)
            .Parameters("parPriceGroupID") = PriceGroupID
            .Parameters.Append .CreateParameter("parPrice", adCurrency, adParamInput)
            .Parameters("parPrice") = NewPrice
            'added 09.25.2004
            .Parameters.Append .CreateParameter("parQtyPack", adInteger, adParamInput)
            .Parameters("parQtyPack") = QtyPack
            .Parameters.Append .CreateParameter("parQtyPackPrice", adCurrency, adParamInput)
            .Parameters("parQtyPackPrice") = QtyPackPrice
            .Parameters.Append .CreateParameter("parCost", adCurrency, adParamInput)
            .Parameters("parCost") = Cost

            .Execute

        End With
End Sub


Public Sub ADOXCreateTable()


' The ADOX Catalog object corresponds to the DAO Database


' object
Dim cat As ADOX.Catalog
Dim tbl As ADOX.Table

   ' Open the catalog
   cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=<full path to your database>\<your database name>.mdb;"

   ' Create a new Table object.
   With tbl
      .Name = "Old Invoices"
      ' Create a field and append it to the new Table
      ' object. This must be done before appending the
      ' Table object to the Tables collection of the
      ' Catalog.
      .Columns.Append "OrderID", adVarWChar
   End With

   ' Add the new table to the database.
   cat.Tables.Append tbl


'Clean up the references
   Set cat = Nothing
   Set tbl = Nothing

End Sub

Public Function SetOptions(Confirm As Boolean)
        Application.SetOption "Confirm Record changes", Confirm
        Application.SetOption "Confirm Document deletions", Confirm
        Application.SetOption "Confirm Action queries", Confirm

End Function


Public Function L2N(Letter As String) As Integer
    If gcfHandleErrors Then On Error GoTo ErrLog
    Select Case Letter
    Case "A"
        L2N = 1
    Case "B"
        L2N = 2
    Case "C"
        L2N = 3
    Case "D"
        L2N = 4
    Case "E"
        L2N = 5
    Case "F"
        L2N = 6
    Case "G"
        L2N = 7
    Case "H"
        L2N = 8
    Case Else
        L2N = 0
    End Select
ExitFunc:
    Exit Function
ErrLog:
    L2N = 0
    Resume ExitFunc
End Function


Public Function OnlineBackup()
Dim stOutputName As String
Dim strFTP As String
Dim strTemp As String
Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject

Dim strPassString

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


    If gcfHandleErrors Then On Error GoTo ErrLog

    If glnAgentNumber = 0 Then LoadDefaults
    If glnAgentNumber < 1000 Then glnAgentNumber = glnAgentNumber + 1000

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "Select CashsysInsLoc from Sysfile"
        .CommandType = adCmdText
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
        'If rsTemp.BOF Then
            rstemp.MoveFirst
            strPassString = "c:\kmds\KMDSOLB.bat " & rstemp!CashsysInsLoc & " " & glnAgentNumber & " " & glnVersionNumber
        'End If
    End With
    rstemp.Close
    Set rstemp = Nothing
    Set comQDF = Nothing



    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL")
    End If
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\OUT")
    End If
        If Not fsoTemp.FolderExists("k:\mail\Out\Log") Then
        fsoTemp.CreateFolder ("k:\mail\Out\Log")
    End If

    If Not gcfHandleErrors Then Debug.Print strPassString
    Call Shell(strPassString, vbNormalFocus)

    'End
    'DoCmd.Quit

ExitFunction:
    Exit Function


ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "cmdKMDSFTP_Click"
    Resume ExitFunction

End Function

Public Sub OpenItemForm(ItemID As String)
If gcfHandleErrors Then On Error GoTo Err_cmdOpenItem_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Item"

    stLinkCriteria = "[ID]=" & ItemID
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog

Exit_cmdOpenItem_Click:
    Exit Sub

Err_cmdOpenItem_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpenItem_Click

End Sub

Function MakeUSDate(DateIn As Variant) As String

    ' Do nothing if the value is not a date.
    If Not IsDate(DateIn) Then Exit Function

    ' Convert the date to a U.S. Date format.
    MakeUSDate = "#" & Month(DateIn) & "/" & Day(DateIn) & "/" & Year(DateIn) & "#"
End Function

Function SetRunMailChangesTrue()
    gbRunMailChanges = True

End Function
