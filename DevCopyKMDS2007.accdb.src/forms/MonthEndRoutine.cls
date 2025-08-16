Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub chkFTPFZActive_Click()
    SetKMDSSettings "FTPFZActive", chkFTPFZActive
    cmdGetPriceFile.Visible = chkFTPFZActive
End Sub

Private Sub cmdClose_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdClose_Click
    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
End Sub

Private Sub cmdGetPriceFile_Click()
Dim strFZDir As String
Dim strFZExe As String
Dim stAppName As String

Dim fsoTemp As New FileSystemObject
    If gcfHandleErrors Then On Error GoTo Err_cmdGetPriceFile_Click
    If chkFTPFZActive Then
'    txtstrtemp.WriteLine ("get newprice.txt newprice.fil")
'    txtstrtemp.WriteLine ("get full_pricelist.txt fulprice.txt")
        If fsoTemp.FileExists("c:\npos\NEWPRICE_ICExpansion.fil") Then
            fsoTemp.DeleteFile ("c:\npos\NEWPRICE_ICExpansion.fil")
        End If
        If fsoTemp.FileExists("c:\npos\NEWPRICE_ICExpansion.txt") Then
            fsoTemp.DeleteFile ("c:\npos\NEWPRICE_ICExpansion.txt")
        End If
        If fsoTemp.FileExists("C:\npos\full_pricelist_ICExpansion.txt") Then
            fsoTemp.DeleteFile ("C:\npos\full_pricelist_ICExpansion.txt")
        End If


        If fsoTemp.FileExists("C:\npos\Price File\NEWPRICE_ICExpansion.txt") Then
            If Not fsoTemp.FolderExists("C:\npos\Price File") Then
                fsoTemp.CreateFolder ("C:\npos\Price File")
            End If
            If Not fsoTemp.FolderExists("C:\npos\Price File\Archive") Then
                fsoTemp.CreateFolder ("C:\npos\Price File\Archive")
            End If
            fsoTemp.MoveFile "C:\npos\Price File\NEWPRICE_ICExpansion.txt", "C:\npos\Price File\Archive\np" & Format(Now, "yy-mm-dd-hh-nn") & ".txt"
        End If

        If fsoTemp.FileExists("C:\npos\Price File\full_pricelist_ICExpansion.txt") Then
            fsoTemp.DeleteFile ("C:\npos\Price File\full_pricelist_ICExpansion.txt")
        End If

        strFZDir = GetKMDSSettings("FTPFZDir", "C:\Program Files (x86)\FileZilla FTP Client\") 'false
        If strFZDir = "C:\Program Files\FileZilla FTP Client\" Then
            SetKMDSSettings "FTPFZDir", "C:\Program Files (x86)\FileZilla FTP Client\"
        End If
        strFZDir = GetKMDSSettings("FTPFZDir", "C:\Program Files (x86)\FileZilla FTP Client\") 'false
        strFZExe = GetKMDSSettings("FTPFZExe", "FileZilla.exe") 'false
        strFZDir = Trim(strFZDir)
        strFZExe = strFZExe
        stAppName = strFZDir & strFZExe & " --site=" & Chr(34) & "0/OLCC FTP"
        stAppName = stAppName & "/Price File" & Chr(34)
        'MsgBox "Using Filezilla to receive Price File" & vbCrLf & vbCrLf & "Choose 'Bookmarks -> Price File" & vbCrLf & "Then double-click on 2 files: newprice.txt AND full_pricelist.txt  on the right"
        MsgBox "Using Filezilla to receive Price File" & vbCrLf & vbCrLf & "Double-click on 2 files on the RIGHT:" & vbCrLf & vbCrLf & "NEWPRICE_ICExpansion.txt " & vbCrLf & "full_pricelist_ICExpansion.txt"
        'Call Shell(stAppName, 1)
        WaitForExecCmd (stAppName)
        fsoTemp.CopyFile "C:\NPOS\Price file\NEWPRICE_ICExpansion.txt", "C:\npos\NEWPRICE_ICExpansion.txt", True
        fsoTemp.CopyFile "C:\NPOS\Price file\full_pricelist_ICExpansion.txt", "C:\npos\full_pricelist_ICExpansion.txt", True
    Else
        'do nothing
        'download price files thru ftp script
    End If

Exit_cmdGetPriceFile_Click:
    Exit Sub

Err_cmdGetPriceFile_Click:
    MsgBox Err.Description
    Resume Exit_cmdGetPriceFile_Click
End Sub

Private Sub Form_Load()
    chkFTPFZActive = GetKMDSSettings("FTPFZActive", "0") 'false
    cmdGetPriceFile.Visible = chkFTPFZActive
'    added 31415
    chkUseWinSCPForOLASTransmits = GetKMDSSettings("UseWinSCPForOLASTransmits", "False") 'false
End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
    If Not Cancel Then
        If MsgBox("This will apply the New Prices" & vbCrLf & vbCrLf & "Do you want to continue?", vbYesNo, "Month End Routine") = vbNo Then
            Cancel = True
        End If
    End If
End Sub


Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub

Private Sub cmdExecute_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdExecute_Click


    EndOFMonth
    'added 10/3/2004
    MailChanges (False)

    'added 2010-10-12
    GetOLCCItemFiles
'    DoCmd.OpenReport "70 Web UPC Code Discrepancy", acViewNormal, "", "", acNormal

    MsgBox "Complete!"

Exit_cmdExecute_Click:
    DoCmd.Close acForm, Me.Name
    Exit Sub

Err_cmdExecute_Click:
    MsgBox Err.Description
    Resume Exit_cmdExecute_Click

End Sub
