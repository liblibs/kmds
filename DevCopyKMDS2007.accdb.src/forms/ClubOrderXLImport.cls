Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim fsoTemp As New FileSystemObject

Private Sub Command2_Click()

Dim strTemp As String
Dim strDestError As String

    strTemp = Dir("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN\*.xls")

    Do While strTemp <> vbNullString
        txtFileName = ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN\" & strTemp)
        ImportClubOrderFromExcel txtFileName
        If fsoTemp.FileExists(txtFileName) Then  'IT DIDN'T GET MOVED... MOVE IT TO ERROR FILE
            strDestError = Replace(txtFileName, "\IN\", "\IN\ERROR\" & Format(Now, "yyyymmdd hh nn ss"))
            fsoTemp.MoveFile txtFileName, strDestError
        End If
        strTemp = Dir
    Loop
    MsgBox "Finished processing; The orders must be printed to recalculate the totals", vbCritical

    DoCmd.Close

    DoCmd.OpenForm "ClubOrderCreate"

End Sub


Private Sub Form_Load()
Me.InsideHeight = boxScreen.Height + 500
Me.InsideWidth = boxScreen.Width + 500
End Sub

Private Sub Form_Open(Cancel As Integer)
Dim strTemp As String

    If gstrCashsysInsLoc = "" Then LoadDefaults
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN")
    End If
    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN\ARCHIVE") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN\ARCHIVE")
    End If

    If Not fsoTemp.FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN\ERROR") Then
        fsoTemp.CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN\ERROR")
    End If

    strTemp = Dir("\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN\*.xls")
    If strTemp = vbNullString Then
        MsgBox "No excel files found in " & vbCrLf & "\\" & gstrCashsysInsLoc & "\KMDS\MAIL\IN\*.xls"
        Cancel = True
        Exit Sub
    End If

    Do While strTemp <> vbNullString
        Text0 = Text0 & strTemp & vbCrLf
        strTemp = Dir
    Loop

End Sub


Private Sub ImportClubOrderFromExcel(SpreadsheetFile As String)

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")


    Dim nOrderNumber As Long
    Dim strDest As String
    strDest = Replace(SpreadsheetFile, "\IN\", "\IN\ARCHIVE\" & Format(Now, "yyyymmdd hh nn ss"))

    If gcfHandleErrors Then On Error GoTo ErrLog
'    DoCmd.comQDF ("Delete * from ClubOrderXL")
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "Delete * from ClubOrderXL"
        .CommandType = adCmdText
        .Execute
    End With

    Set comQDF = Nothing
    Set comQDF = CreateObject("ADODB.COMMAND")

    DoCmd.TransferSpreadsheet , acSpreadsheetTypeExcel9, "ClubOrderXL", SpreadsheetFile, True

    Me.Requery ' to get the club id
    Me.Refresh

    nOrderNumber = CreateClubOrder(ClubID, False, ClubID)
    txtOrderNumber = nOrderNumber
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryClubOrderXLAppend"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parOrderNumber", adBigInt, adParamInput)
        .Parameters("parOrderNumber") = nOrderNumber
        .Execute
    End With
    Set comQDF = Nothing

    fsoTemp.MoveFile SpreadsheetFile, strDest

ExitSub:
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub
