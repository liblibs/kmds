Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim mbEOProcessError As Boolean

Private Sub Command6_Click()
    ProcessCSVFiles (NetworkMailDirFileSpec)
    DoCmd.Close
End Sub
Private Sub ProcessCSVFiles(MailDirFileSpec As String)

    Dim strFile As String

    Dim fsoTemp As FileSystemObject
    Dim fileTemp As File
    Dim filesTemp As Files

    Dim folderTemp As Folder
    Dim S As String
    Dim varReturn As Variant
    Dim strArchiveFileName As String


    If gcfHandleErrors Then On Error GoTo ErrLog
    SetOptions (False)
    Set fsoTemp = CreateObject("Scripting.FileSystemObject")

    'did last end of day finish successfully?
    If fsoTemp.FileExists(NetworkMailDir & "mail\CR1.MDB") Then
        MsgBox NetworkMailDir & "mail\CR1.MDB exists; can not continue..."
        EOLog NetworkMailDir & "mail\CR1.MDB exists; can not continue..."
        mbEOProcessError = True
        Exit Sub
    End If
    If fsoTemp.FileExists(NetworkMailDir & "mail\CR2.MDB") Then
        MsgBox NetworkMailDir & "mail\CR2.MDB exists; can not continue..."
        EOLog NetworkMailDir & "mail\CR2.MDB exists; can not continue..."
        mbEOProcessError = True
        Exit Sub
    End If
    If fsoTemp.FileExists(NetworkMailDir & "mail\CR3.MDB") Then
        MsgBox NetworkMailDir & "mail\CR3.MDB exists; can not continue..."
        EOLog NetworkMailDir & "mail\CR3.MDB exists; can not continue..."
        mbEOProcessError = True
        Exit Sub
    End If
    If fsoTemp.FileExists(NetworkMailDir & "mail\CR4.MDB") Then
        MsgBox NetworkMailDir & "mail\CR4.MDB exists; can not continue..."
        EOLog NetworkMailDir & "mail\CR4.MDB exists; can not continue..."
        mbEOProcessError = True
        Exit Sub
    End If
    'yes

    'copy shell files to begin end of day
    fsoTemp.CopyFile "C:\KMDS\CashRegisterDetailTemplate.mdb", NetworkMailDir & "mail\CR1.MDB"
    fsoTemp.CopyFile "C:\KMDS\CashRegisterDetailTemplate.mdb", NetworkMailDir & "mail\CR2.MDB"
    fsoTemp.CopyFile "C:\KMDS\CashRegisterDetailTemplate.mdb", NetworkMailDir & "mail\CR3.MDB"
    fsoTemp.CopyFile "C:\KMDS\CashRegisterDetailTemplate.mdb", NetworkMailDir & "mail\CR4.MDB"

    If Not fsoTemp.FolderExists(MailDirFileSpec) Then
        fsoTemp.CreateFolder (MailDirFileSpec)
    End If
    If Not fsoTemp.FolderExists(MailDirFileSpec & "Archive") Then
        fsoTemp.CreateFolder (MailDirFileSpec & "Archive")
    End If


    Set folderTemp = fsoTemp.GetFolder(MailDirFileSpec)
    Set filesTemp = folderTemp.Files

    For Each fileTemp In filesTemp
        If InStr(fileTemp.Name, ".csv") = (Len(fileTemp.Name) - 3) Then
            S = S & "'" & fileTemp.Path & "';"
            txtProcessed.RowSource = S
            EOLog Me.Name & " importing " & fileTemp.Path
            DoCmd.TransferText acImportDelim, "InvoiceLineLDImportSpec", "InvoLine4", fileTemp.Path, False
            strArchiveFileName = MailDirFileSpec & "Archive\" & fileTemp.Name
            'duplicate file names in archive directory not allowed
            If fsoTemp.FileExists(strArchiveFileName) Then
                strArchiveFileName = strArchiveFileName & Format(Now, "yyyymmddhhnn")
            End If
            fsoTemp.MoveFile fileTemp.Path, strArchiveFileName
        End If
    Next

    'trantype is required in involine
    EOLog "qryInvoLine4UpdateNullTransTypesToS"
    CurrentDb().Execute "qryInvoLine4UpdateNullTransTypesToS"
    'for importing new items
    EOLog "qryInvoLine4UpdateDrawerToSizeID"
    CurrentDb().Execute "qryInvoLine4UpdateDrawerToSizeID"
    'append new items
    EOLog "qryInvoLine4AppendNewItems"
    CurrentDb().Execute "qryInvoLine4AppendNewItems"
    txtToBeProcessed.RowSource = ""
    Me.Repaint
    'let's do the end of day
    EOLog "Starting EndOFDayStream"

    EndOFDayStream

ExitSub:
    SetOptions (True)
    Exit Sub
ErrLog:
    EOLog Me.Name & " ERROR IMPORTING PROCESSCSVFILES " & Err.Number & " " & Err.Description
    EMsgBox Me, "ProcessCSVFiles" & vbCrLf & Err.Number & " " & Err.Description
    mbEOProcessError = True
    Resume Next
End Sub

Private Sub Form_Close()
Dim strMsg As String
    If mbEOProcessError Then
        strMsg = gstrMailDir & "Log\"
        strMsg = strMsg & Format(Now, "yyyy") & "\"
        strMsg = strMsg & Format(Now, "mm") & "\"
        strMsg = strMsg & Format(Now, "dd") & "\"
        strMsg = strMsg & "EOLog.Txt"
        StartDoc strMsg
    End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    EMsgBox Me, "Form_Error " & DataErr & vbCrLf & Err.Number & " " & Err.Description
    mbEOProcessError = True
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim fso As FileSystemObject
    Dim fileTemp As File
    Dim filesTemp As Files

    Dim folderTemp As Folder
    Dim S As String

    'use register #4 file to import the POS

    If gcfHandleErrors Then On Error GoTo ErrLog
    mbEOProcessError = False
    EOLog "Opening " & Me.Name

    EOLog "Mail dir: " & NetworkMailDir & "mail\in\"
    NetworkMailDirFileSpec = NetworkMailDir & "mail\in\"

    Set fso = CreateObject("Scripting.FileSystemObject")


    Set folderTemp = fso.GetFolder(NetworkMailDirFileSpec)
    Set filesTemp = folderTemp.Files
    For Each fileTemp In filesTemp
        If InStr(fileTemp.Name, ".csv") = (Len(fileTemp.Name) - 3) Then
            S = S & "'" & fileTemp.Path & "';"
        End If
    Next
    If S = "" Then
        EOLog "Nothing to do! No mail in " & vbCrLf & NetworkMailDirFileSpec
        MsgBox "Nothing to do! No mail in " & vbCrLf & NetworkMailDirFileSpec, vbInformation, Me.Name
        Cancel = True
    Else
        EOLog "To be processed: " & S
        txtToBeProcessed.RowSource = S
    End If
ExitSub:
    Exit Sub
ErrLog:
    EOLog "Error: Form_Open :" & Err.Number & " " & Err.Description
    EMsgBox Me, "Form_Open" & vbCrLf & Err.Number & " " & Err.Description
    Cancel = True
    mbEOProcessError = True
    Resume ExitSub
End Sub
