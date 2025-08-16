Attribute VB_Name = "modPSC"
Option Compare Database
Option Explicit

Public Function Dscrptn(Description As String) As String
    Dscrptn = Replace(Description, "a", "")
    Dscrptn = Replace(Dscrptn, "e", "")
    Dscrptn = Replace(Dscrptn, "i", "")
    Dscrptn = Replace(Dscrptn, "o", "")
    Dscrptn = Replace(Dscrptn, "u", "")
    Dscrptn = Left(Dscrptn, 10)
    Dscrptn = Format(Dscrptn, "!@@@@@@@@@@")

End Function


Public Sub CreatePSCLookupFiles()
    DoCmd.TransferText acExportDelim, "QryPSCDownloadItem Export Specification", "qryPSCDownloadItem", "C:\PDTFILES\Item.txt", False
    DoCmd.TransferText acExportDelim, "QryPSCDownloadDescr Export Specification", "qryPSCDownloadDescr", "C:\PDTFILES\Descr.txt", False
    MsgBox "Files created... From the main menu on the PSC Press F3 F2"
    Shell "C:\PDTFiles\PSCItemFile.bat", vbNormalFocus
End Sub


Public Sub CreatePSCImportScannedData()

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

Dim strMsg As String

    If gcfHandleErrors Then On Error GoTo ErrLog



    SetOptions False

    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryPSCDownloadRecordCount"
        .CommandType = adCmdStoredProc
    End With
    With rstemp
        .Open comQDF, , adOpenDynamic, adLockOptimistic
        If rstemp!records > 0 Then
            strMsg = "There are " & rstemp!records & " records from a previous download." & vbCrLf & vbCrLf
            strMsg = strMsg & "Do you want to delete them?"
            If MsgBox(strMsg, vbYesNo) = vbYes Then
                DoCmd.OpenQuery "qryPSCInput_Step1_DeleteRecords"
            End If
        End If
        .Close
    End With
    Set rstemp = Nothing
    Set comQDF = Nothing



    WaitForExecCmd (CStr("c:\pdtfiles\PSCInput.bat"))
    DoCmd.TransferText , "PSCINPUT Import Specification", "PSCInput", "c:\pdtFiles\PSCInput.txt", False
    DoCmd.OpenQuery "qryPSCInput_Step2_UpdateWithItemInfo"
    SetOptions True
    ArchivePSCDownload
    DoCmd.OpenForm "PSCInputEditScans"

ExitSub:
    Exit Sub


'503-697-7251
ErrLog:
    MsgBox "CreatePSCImportScannedData" & vbCrLf & Err.Number & " " & Err.Description
    Resume ExitSub

End Sub


Public Sub ArchivePSCDownload()

Dim fso As FileSystemObject, msg, f
Dim fldr As String
    If gcfHandleErrors Then On Error GoTo ErrLog
    If gstrReportsDir = "" Then LoadDefaults

    'create the directory for today e.g. "c:\KMDS\reports\2003\12\25"
    Set fso = CreateObject("Scripting.FileSystemObject")


    fldr = gstrReportsDir
    If Not (fso.FolderExists(fldr)) Then
          Set f = fso.CreateFolder(fldr)
    End If

    fldr = fldr & "PSC" & "\"
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
    fldr = fldr & "PSCInput"

    fldr = fldr & Format(Now, "yyyy-mm-dd hh-nn-ss") & ".txt"
    fso.MoveFile "c:\pdtFiles\PSCInput.txt", fldr

ExitSub:
    Set fso = Nothing
    Exit Sub

ErrLog:
    MsgBox "ArchivePSCDownload" & vbCrLf & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub
