Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Audit Item Reconciliation File Format
'
'Filename: audit.txt  - header record
'Field Name     Data Type           Description
'Creation Date   Numeric(6,0)    Format YYMMDD. Date file was created by OLCC audit staff.
'Agency Number   Numeric(4,0)    Agency number to which these quantity differences apply.
'Record Count    Numeric(3,0)    Record count for file.
'
'Filename: audit.txt  - detail record
'Field Name                  Data Type       Description
'OLCC Numeric Item Code      Numeric(6,0)    Zero filled to the left.
'Item Quantity Difference    Numeric(5,0)    Computed difference in quantity between OLCC audit item count and POS item count as provided by POS system: (OLCC Item Count) - (POS Item Count) = Item Quantity Difference.  Zero filled to the left.  Negative values include a minus sign (-) in the first position: e.g. -0001.
'1301311138004
'00140200005
'001406-0001

Private Sub Form_Load()
Dim fsoTemp As New FileSystemObject
Dim nI As Integer
Dim nasc As Integer

If gcfHandleErrors Then On Error GoTo ErrLog
    txtDriveLetter = GetKMDSSettings("AuditFileDrive", "F:")
    SetAuditTextBoxes
    Form.InsideHeight = boxScreen.Height + 500 '+ cmdGoToFirstRecord.Height
    Form.InsideWidth = boxScreen.Width + 500
    nasc = Asc("A")
    For nI = 0 To 25
        If fsoTemp.DriveExists(Chr(nasc + nI)) Then
            If fsoTemp.Drives.Item(Chr(nasc + nI)).DriveType = Removable Then
                lstDrives.AddItem fsoTemp.GetDrive(Chr(nasc + nI))
            End If
        End If
    Next nI
    For nI = 0 To 25
        If fsoTemp.DriveExists(Chr(nasc + nI)) Then
            If fsoTemp.Drives.Item(Chr(nasc + nI)).DriveType <> Removable Then
                lstDrivesFixed.AddItem fsoTemp.GetDrive(Chr(nasc + nI))
            End If
        End If
    Next nI

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "Form_Load"
    Resume ExitSub
End Sub


Private Sub lstDrives_Click()
    txtDriveLetter = lstDrives
    SetAuditTextBoxes
End Sub

Private Sub SetAuditTextBoxes()
    Dim fsoTemp As New FileSystemObject
    Dim fileTemp As File
    Dim strSource As String
    Dim strTarget As String
    On Error GoTo ErrFile
    lblDateCreated.Caption = "File not found."

    txtAuditFile = Trim(txtDriveLetter) & "\Audit.txt"
    Set fileTemp = fsoTemp.GetFile(txtAuditFile)
    lblDateCreated.Caption = fileTemp.DateLastModified
    If Dir(txtAuditFile) <> "" Then
        cmdCreateAuditBatch.Visible = True
        ' added 10/19/2023
        strSource = txtAuditFile
        strTarget = "C:\NPOS\AuditFile\Audit-" & Format(Now, "yymmddhhnn") & ".txt"
        If Not fsoTemp.FolderExists("C:\NPOS") Then
            fsoTemp.CreateFolder ("C:\NPOS")
        End If
        If Not fsoTemp.FolderExists("C:\NPOS\AuditFile") Then
            fsoTemp.CreateFolder ("C:\NPOS\AuditFile")
        End If
        fsoTemp.CopyFile strSource, strTarget, True
        ' 10/19/2023
    Else
        MsgBox txtAuditFile & " not found!"
        cmdCreateAuditBatch.Visible = False
    End If

ExitSub:
    lblDateCreated.Visible = True
    lblCreateAuditBatch.Visible = Not cmdCreateAuditBatch.Visible
    Exit Sub
ErrFile:
    cmdCreateAuditBatch.Visible = False
    Resume ExitSub
End Sub

Private Sub lstDrivesFixed_Click()
    txtDriveLetter = lstDrivesFixed
    SetAuditTextBoxes
End Sub
'------------------------------------------------------------
' cmdCreateAuditBatch_Click
'
'------------------------------------------------------------
Private Sub cmdCreateAuditBatch_Click()
   If gcfHandleErrors Then On Error GoTo cmdCreateAuditBatch_Click_Err


   'Dim rstemp As New ADODB.Recordset
    Dim rstemp As Object
    Dim strSource  As String
    Dim strTarget As String
    Dim fsoTemp As New FileSystemObject

    Set rstemp = CreateObject("ADODB.Recordset")


    If Abs(DateDiff("D", Now, lblDateCreated.Caption)) > 3 Then
        txtAuditFile.BackColor = vbRed
        lblDateCreated.ForeColor = vbRed
        Me.Refresh
        Me.Repaint
        If MsgBox("The audit file is more than 3 days old!  Are you sure!", vbYesNo, lblDateCreated.Caption) = vbNo Then
            GoTo cmdCreateAuditBatch_Click_Exit
        End If
    End If


    SetOptions False

    rstemp.Open "Audit", CurrentProject.Connection, adOpenDynamic, adLockOptimistic


    If rstemp.EOF And rstemp.BOF Then
'        DoCmd.OpenQuery "qryScanBatchAppendToAudit", acViewNormal, acEdit
'        DoCmd.OpenQuery "qryScanBatchAppendToAuditRecalcVariance", acViewNormal, acEdit
'        DoCmd.OpenQuery "qryScanBatchAppendToInvoline", acViewNormal, acEdit
'        ExportScanBatchLineCounts
'        DoCmd.RunSQL "Delete * from ScanBatchLineCounts Where ScanBatchHeaderID = 3" '& txtScanBatchHeaderId
'        DoCmd.Close acForm, "ScanBatch"

        SetKMDSSettings "AuditFileDrive", txtDriveLetter
        DoCmd.RunSQL ("Delete * from OLCCAuditVariance")
        DoCmd.TransferText acImportFixed, "OLCCAuditVarianceImportSpec", "OLCCAuditVariance", txtAuditFile, False
        DoCmd.OpenQuery "qryOLCCAuditVarianceAppendToAudit", acViewNormal, acEdit
        ' added 10/19/2023
        strSource = txtAuditFile
        strTarget = "C:\NPOS\AuditFile\Audit-" & Format(Now, "yymmddhhnn") & "-D7Created.txt"
        If Not fsoTemp.FolderExists("C:\NPOS") Then
            fsoTemp.CreateFolder ("C:\NPOS")
        End If
        If Not fsoTemp.FolderExists("C:\NPOS\AuditFile") Then
            fsoTemp.CreateFolder ("C:\NPOS\AuditFile")
        End If
        fsoTemp.CopyFile strSource, strTarget, True
        ' 10/19/2023
        MsgBox "Variance counts have been added" & vbCrLf & "Please post these counts in Menu Option D-7"
        DoCmd.OpenForm "Audit"
        DoCmd.Close acForm, Me.Name
   Else
        'StatusLine "The D-7 Audit Inventory screen already has a batch pending; this batch must be posted or deleted before this is allowed!", vbRed, True
        MsgBox "The D-7 Audit Inventory screen already has a batch pending; this batch must be posted or deleted before this is allowed!", vbCritical
    End If
cmdCreateAuditBatch_Click_Exit:
    Exit Sub

cmdCreateAuditBatch_Click_Err:
    MsgBox Error$
    Resume cmdCreateAuditBatch_Click_Exit

End Sub
