Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'------------------------------------------------------------
' cmdDriizly_Click



'•Store Location ID - the internal store number identifier
'•SKU/Product ID - the internal identifier for each item in the store.
'•Name/Description - what the item is called
'•Retail Price(s) – in store retail
'•Sale/Promo Pricing - sale price include start & end dates if available
'•Quantity on Hand - what the store has left to sell
'•UPC - universal product code (Barcode)
'•Alt(s) UPC - additional UPCs for the same item
'•Sales Tax Code(s) – system tax code(s)
'•Size - size of bottle/can,
'•Pack Qty – 4pk, 6pk etc.
'•Pack price – retail price & sale/promo price
'•Container -bottle Or can
'•Vintage - year of wine if stored separately
'•Item Status – inactive and or deleted items, need identifier within software
'•Other categorical info – i.e.; grape, vineyard, etc.
'•Any other information you think would be helpful, such as bottle deposit etc.

'The test file can be sent to me via email for review.

'Once the file has been approved the file can then be sent to our FTP site, hourly (if possible).
'
'File type -  .csv, .txt or .xml format
'
'File name format  - Drizly Store id_yyyymmdd_0000_inv.csv  (0000 = military time)
'          Example -  4038_20180619_1505_inv.csv
'
'FTP Credentials -
'
'Rose City Liquor (OR)
'Drizly Store ID =4038
'FTP Host = ftp02.drzscs.com
'UserName = rosecityliquor
'Password =6wFQa975gXfJ5ar9
'------------------------------------------------------------
Private Sub cmdDriizly_Click()
On Error GoTo cmdDriizly_Click_Err
    SetOptions False
    DoCmd.OpenQuery "qryDrizly_01_Delete", acViewNormal, acEdit
    DoCmd.OpenQuery "qryDrizly_02_Append", acViewNormal, acEdit
    DoCmd.TransferText acExportDelim, , "qryDrizly_03_Select", txtFileName, True
   'doCmd.TransferText acExportDelim, , "qryFintechExport", strFileName, True


    SetOptions True
    DoCmd.OpenTable "Drizly"

cmdDriizly_Click_Exit:
    Exit Sub

cmdDriizly_Click_Err:
    MsgBox Error$
    Resume cmdDriizly_Click_Exit

End Sub


'------------------------------------------------------------
' Command1_Click
'
'------------------------------------------------------------
Private Sub Command1_Click()
On Error GoTo Command1_Click_Err

    On Error Resume Next
    DoCmd.GoToControl Screen.PreviousControl.Name
    DoCmd.RunCommand acCmdAutoDial


Command1_Click_Exit:
    Exit Sub

Command1_Click_Err:
    MsgBox Error$
    Resume Command1_Click_Exit

End Sub


'------------------------------------------------------------
' cmdSaveSettings_Click
'
'------------------------------------------------------------
Private Sub cmdSaveSettings_Click()
On Error GoTo cmdSaveSettings_Click_Err

    SetKMDSSettings "DrizlyStoreID", txtDrizlyStoreID
    SetKMDSSettings "DrizlyFTPHost", txtDrizlyFTPHost
    SetKMDSSettings "DrizlyUsername", txtDrizlyUserName
    SetKMDSSettings "DrizlyPassword", txtDrizlyPassword




cmdSaveSettings_Click_Exit:
    Exit Sub

cmdSaveSettings_Click_Err:
    MsgBox Error$
    Resume cmdSaveSettings_Click_Exit

End Sub


Private Sub Form_Open(Cancel As Integer)
Dim bDIInstalled As Boolean

    diinstalled = GetKMDSSettings("DIInstalled", "False")
    If Not diinstalled Then
        MsgBox "Drizly interface not installed! Contact KMDS for Info"
        Cancel = True
    Else
        txtDrizlyStoreID = GetKMDSSettings("DrizlyStoreID", "n/a")
        txtDrizlyFTPHost = GetKMDSSettings("DrizlyFTPHost", "n/a")
        txtDrizlyUserName = GetKMDSSettings("DrizlyUsername", "n/a")
        txtDrizlyPassword = GetKMDSSettings("DrizlyPassword", "n/a")
        'File name format  - Drizly Store id_yyyymmdd_0000_inv.csv  (0000 = military time)
'          Example -  4038_20180619_1505_inv.csv
'
        txtFileName = "C:\kmds\Mail\Out\" & txtDrizlyStoreID & Format(Now, "_yyyymmdd_hhnn_") & "inv.csv"
    End If


End Sub
