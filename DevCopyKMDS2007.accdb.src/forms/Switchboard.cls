Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdQuitters_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdQuitters_Click
    DoCmd.Quit

Exit_cmdQuitters_Click:
    Exit Sub

Err_cmdQuitters_Click:
    MsgBox Err.Description
    Resume Exit_cmdQuitters_Click

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not gcfHandleErrors Then Debug.Print KeyCode; Chr(KeyCode)
    Select Case KeyCode
    Case 65 'A
        ProcessKeyStroke (Chr(KeyCode))
    Case 66 'B
        ProcessKeyStroke (Chr(KeyCode))
    Case 67 'C
        ProcessKeyStroke (Chr(KeyCode))
    Case 68 'D
        ProcessKeyStroke (Chr(KeyCode))
    Case 69 'E
        ProcessKeyStroke (Chr(KeyCode))
    Case 70 'F
        ProcessKeyStroke (Chr(KeyCode))
    Case 71 'G
        ProcessKeyStroke (Chr(KeyCode))
    Case 72 'H
        ProcessKeyStroke (Chr(KeyCode))
    Case 49
        ProcessKeyStroke (Chr(KeyCode))
    Case 50
        ProcessKeyStroke (Chr(KeyCode))
    Case 51
        ProcessKeyStroke (Chr(KeyCode))
    Case 52
        ProcessKeyStroke (Chr(KeyCode))
    Case 53
        ProcessKeyStroke (Chr(KeyCode))
    Case 54
        ProcessKeyStroke (Chr(KeyCode))
    Case 55
        ProcessKeyStroke (Chr(KeyCode))
    Case 56
        ProcessKeyStroke (Chr(KeyCode))

    Case 97 'a = 1 on the 10 keypad!
        ProcessKeyStroke ("1")
    Case 98  ' b= 2on the 10 keypad!
        ProcessKeyStroke ("2")
    Case 99  ' c = 3 on the 10 keypad!
        ProcessKeyStroke ("3")
    Case 100 ' d = 4 on the 10 keypad!
        ProcessKeyStroke ("4")
    Case 101 ' e = 5 on the 10 keypad!
        ProcessKeyStroke ("5")
    Case 102 ' f = 6 on the 10 keypad!
        ProcessKeyStroke ("6")
    Case 103 ' g = 7 on the 10 keypad!
        ProcessKeyStroke ("7")
    Case 104 ' h = 8 on the 10 keypad!
        ProcessKeyStroke ("8")

    Case Else
    End Select
End Sub

Private Sub ProcessKeyStroke(CharStroke As String)
    If Left(OptionLabel2.Caption, 1) = CharStroke Then
        HandleButtonClick (2)
    End If

    If Left(OptionLabel3.Caption, 1) = CharStroke Then
        HandleButtonClick (3)
    End If
    If Left(OptionLabel4.Caption, 1) = CharStroke Then
        HandleButtonClick (4)
    End If
    If Left(OptionLabel5.Caption, 1) = CharStroke Then
        HandleButtonClick (5)
    End If
    If Left(OptionLabel6.Caption, 1) = CharStroke Then
        HandleButtonClick (6)
    End If
    If Left(OptionLabel7.Caption, 1) = CharStroke Then
        HandleButtonClick (7)
    End If
    If Left(OptionLabel8.Caption, 1) = CharStroke Then
        HandleButtonClick (8)
    End If

End Sub


Private Sub Form_Load()
    txtName = gstrStoreName
    lblDebug.Visible = Not gcfHandleErrors
    GetKMDSSettings "BottleDepositQtyAdjAtEOD", "False"
    '    If gcfHandleErrors Then
'        DoCmd.ShowToolbar "Ribbon", acToolbarNo
'    Else
'        DoCmd.ShowToolbar "Ribbon", acToolbarYes
'    End If
'
End Sub

Private Sub Form_Open(Cancel As Integer)
' Minimize the database window and initialize the form.
    Dim fso As FileSystemObject
    Dim ts As TextStream
    Dim ts2 As TextStream
    Dim Quote As String
    Quote = Chr(34)


Dim rs
'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")


'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

    'added 02/09/2005
    If gcfHandleErrors Then On Error GoTo ErrLog
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryCashiersAndProgramAllowedCount"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parCashierID", adInteger, adParamInput)
        .Parameters("parCashierID") = glnCashierID
    End With

    rstemp.Open comQDF, , adOpenStatic, adLockReadOnly
    With rstemp
        If .BOF And .EOF Then
        Else
            If !Count = 1 Then ' run cash register only!
                DoCmd.OpenForm "CashRegister"
                Cancel = True
                Exit Sub
            End If

        End If
        .Close
    End With
    Set rstemp = Nothing
    Set comQDF = Nothing


    ' Move to the switchboard page that is marked as the default.
    Me.Filter = "[ItemNumber] = 0 AND [Argument] = 'Default' "
    Me.FilterOn = True
    Set rs = Me.RecordsetClone
    rs.Bookmark = Me.Bookmark
    glnVersionNumber = rs!Command
    txtVersion = " Ver " & Format(Now, "yyyy") & "." & glnVersionNumber

    rs.Close
    Set rs = Nothing
    If Not gcfHandleErrors Then ' create a version number for this directory
        On Error Resume Next
        If Dir("c:\kmds\BackACCDBToCloud.INS") <> "" Then 'ONLY DO THIS ON KMDS DEVELOPMENT MACHINE!
            MkDir "C:\KMDS\Versions" & CStr(glnVersionNumber)
            Set fso = New FileSystemObject

            Set ts = fso.OpenTextFile("C:\kmds\BackACCDBToCloud.bat", ForWriting, True, TristateFalse)
            With ts
                .WriteLine "@# BackACCDBToCloud.bat created by Form_Switchboard Form_Open Routine"
                .WriteLine "@# BackACCDBToCloud.scr created by Form_Switchboard Form_Open Routine"
                .WriteLine "7z u -r c:\kmds\KMDSBackup" & glnVersionNumber & ".zip kmds2007.accdb"
                .WriteLine "7z u -r c:\kmds\KMDSBackup" & glnVersionNumber & ".zip c:\cr3\*.*"
                MsgBox "7z u    c:\kmds\KMDSBackup" & glnVersionNumber & ".zip c:\kmds\backup\*.q*"
                .WriteLine "7z u    c:\kmds\KMDSBackup" & glnVersionNumber & ".zip c:\kmds\backup\*.q*"

                .WriteLine "7z u -r c:\kmds\versionssync" & glnVersionNumber & ".zip versionssync\*.*"
                '.WriteLine "ftp -s:BackACCDBToCloud.scr> BackACCDBToCloud.log"
                .WriteLine "c:"
                .WriteLine "cd \kmds"
                .WriteLine "xcopy versionssync" & glnVersionNumber & ".zip KMDSFtp /y"
                .WriteLine "xcopy versionssync" & glnVersionNumber & ".zip versionssync.zip /y"
                .WriteLine "xcopy versionssync.zip KMDSFtp /y"
                .WriteLine "xcopy c:\kmds\KMDSBackup" & glnVersionNumber & ".zip KMDSFtpBackup /y"
                .WriteLine Quote & "C:\Program Files (x86)\FileZilla FTP Client\filezilla.exe" & Quote & " --site=" & Quote & "0/kmds.us" & Quote
                .WriteLine Quote & "C:\Program Files (x86)\FileZilla FTP Client\filezilla.exe" & Quote & " --site=" & Quote & "0/kmds.us.Backup" & Quote
'                .WriteLine "Delete BackACCDBToCloud.scr"

                '.WriteLine "Start BackACCDBToCloud.Log"
                .Close
            End With
            Set ts = Nothing

            Set ts2 = fso.OpenTextFile("C:\kmds\BackACCDBToCloud.scr", ForWriting, True, TristateFalse)
            With ts2
                .WriteLine "open ftp.kmds.us"
                .WriteLine "pdxSteve"
                .WriteLine "$K4521066"
                .WriteLine "binary"
                .WriteLine "cd Agents"
                .WriteLine "Send versionssync.Zip"
                .WriteLine "Send versionssync" & glnVersionNumber & ".zip"
                .WriteLine "cd .."
                .WriteLine "cd KMDSBackup"
                .WriteLine "Send KMDSBackup" & glnVersionNumber & ".zip"
                .WriteLine "cd .."
                .WriteLine "Quit"
                .Close
            End With
            Set ts2 = Nothing
            Set fso = Nothing

        End If
    End If

ExitSub:
    Exit Sub
ErrLog:
    EMsgBox Me, "Form_Open " & Err.Number & " " & Err.Description
    Resume ExitSub
End Sub

Private Sub Form_Current()
' Update the caption and fill in the list of options.

    Label1.Caption = Nz(Me![ItemText], "")
    'Label2.Caption = Nz(Me![ItemText], "")

    FillOptions

End Sub

Private Sub FillOptions()
' Fill in the options for this switchboard page.

    ' The number of buttons on the form.
    Const conNumButtons = 8

    Dim con As Object
    Dim rs As Object
    Dim stSql As String
    Dim intOption As Integer
    Dim bShowOption As Boolean


    ' Set the focus to the first button on the form,
    ' and then hide all of the buttons on the form
    ' but the first.  You can't hide the field with the focus.
    Me![Command24].SetFocus
    For intOption = 1 To conNumButtons
        Me("Option" & intOption).Visible = False
        Me("OptionLabel" & intOption).Visible = False
    Next intOption

    ' Open the table of Switchboard Items, and find
    ' the first item for this Switchboard Page.
    Set con = Application.CurrentProject.Connection
    stSql = "SELECT * FROM [Switchboard Items]"
    stSql = stSql & " WHERE [ItemNumber] > 0 AND [SwitchboardID]=" & Me![SwitchboardID]
    stSql = stSql & " ORDER BY [ItemNumber];"

    stSql = "SELECT * "
    stSql = stSql & " FROM [Switchboard Items] "
    stSql = stSql & "  LEFT JOIN CashiersAndPrograms "
    stSql = stSql & "   ON [Switchboard Items].Argument = CashiersAndPrograms.ProgramName "
    stSql = stSql & " WHERE ((([Switchboard Items].ItemNumber) > 0) "
    stSql = stSql & "   And  (([Switchboard Items].SwitchboardID) = " & Me![SwitchboardID] & ") "
    stSql = stSql & "   And   ((CashiersAndPrograms.Allowed) Is Null Or (CashiersAndPrograms.Allowed) <> 0) "
    stSql = stSql & "   And    ((CashiersAndPrograms.CashierID) Is Null Or (CashiersAndPrograms.CashierID) = " & glnCashierID & ")) "
    stSql = stSql & " ORDER BY [Switchboard Items].ItemNumber;"


    Set rs = CreateObject("ADODB.Recordset")
    rs.Open stSql, con, 1   ' 1 = adOpenKeyset

    ' If there are no options for this Switchboard Page,
    ' display a message.  Otherwise, fill the page with the items.
    If (rs.EOF) Then
        Me![OptionLabel1].Caption = "There are no items for this Menu page"
    Else
        While (Not (rs.EOF))
            bShowOption = True
            'pay to get access to menu option F-6
            If rs![ItemText] = "6. Metrologic Batch Scanner" Then
                If GetKMDSSettings("MetrologicSubscription", "False") = "False" Then
                    bShowOption = False
                End If
            End If

            If rs![ItemText] = "5. Import Orders from Website" Then
                If GetKMDSSettings("IOWSub", "False") = "False" Then
                    bShowOption = False
                End If
            End If

            If bShowOption Then
                Me("Option" & rs![ItemNumber]).Visible = True
                Me("OptionLabel" & rs![ItemNumber]).Visible = True
                Me("OptionLabel" & rs![ItemNumber]).Caption = rs![ItemText]
            End If

            rs.MoveNext
        Wend
    End If

    ' Close the recordset and the database.
    rs.Close
    Set rs = Nothing
    Set con = Nothing

End Sub

Private Function HandleButtonClick(intBtn As Integer)
' This function is called when a button is clicked.
' intBtn indicates which button was clicked.

    ' Constants for the commands that can be executed.
    Const conCmdGotoSwitchboard = 1
    Const conCmdOpenFormAdd = 2
    Const conCmdOpenFormBrowse = 3
    Const conCmdOpenReport = 4
    Const conCmdCustomizeSwitchboard = 5
    Const conCmdExitApplication = 6
    Const conCmdRunMacro = 7
    Const conCmdRunCode = 8
    Const conCmdOpenPage = 9

    ' An error that is special cased.
    Const conErrDoCmdCancelled = 2501

    Dim con As Object
    Dim rs As Object
    Dim stSql As String

If gcfHandleErrors Then On Error GoTo HandleButtonClick_Err

    ' Find the item in the Switchboard Items table
    ' that corresponds to the button that was clicked.
    Set con = Application.CurrentProject.Connection
    Set rs = CreateObject("ADODB.Recordset")
    stSql = "SELECT * FROM [Switchboard Items] "
    stSql = stSql & "WHERE [SwitchboardID]=" & Me![SwitchboardID] & " AND [ItemNumber]=" & intBtn
    rs.Open stSql, con, 1    ' 1 = adOpenKeyset

    ' If no item matches, report the error and exit the function.
    If (rs.EOF) Then
        MsgBox "There was an error reading the Switchboard Items table."
        rs.Close
        Set rs = Nothing
        Set con = Nothing
        Exit Function
    End If

    Select Case rs![Command]

        ' Go to another switchboard.
        Case conCmdGotoSwitchboard
            Me.Filter = "[ItemNumber] = 0 AND [SwitchboardID]=" & rs![Argument]

        ' Open a form in Add mode. (datasheet)
        Case conCmdOpenFormAdd
            DoCmd.OpenForm rs![Argument], acFormDS
        ' Open a form.
        Case conCmdOpenFormBrowse
            DoCmd.OpenForm rs![Argument]

        ' Open a report.
        Case conCmdOpenReport
            DoCmd.OpenReport rs![Argument], acPreview

        ' Customize the Switchboard.
        Case conCmdCustomizeSwitchboard
            ' Handle the case where the Switchboard Manager
            ' is not installed (e.g. Minimal Install).
            If gcfHandleErrors Then On Error Resume Next
            Application.Run "ACWZMAIN.sbm_Entry"
            If (Err <> 0) Then MsgBox "Command not available."
            If gcfHandleErrors Then On Error GoTo 0
            ' Update the form.
            Me.Filter = "[ItemNumber] = 0 AND [Argument] = 'Default' "
            Me.Caption = Nz(Me![ItemText], "")
            FillOptions

        ' Exit the application.
        Case conCmdExitApplication
            CloseCurrentDatabase

        ' Run a macro.
        Case conCmdRunMacro
            DoCmd.RunMacro rs![Argument]

        ' Run code.
        Case conCmdRunCode
            Application.Run rs![Argument]

        ' Open a Data Access Page
        Case conCmdOpenPage
            DoCmd.OpenDataAccessPage rs![Argument]

        ' Any other command is unrecognized.
        Case Else
            MsgBox "Unknown option."

    End Select

    ' Close the recordset and the database.
    rs.Close

HandleButtonClick_Exit:
If gcfHandleErrors Then On Error Resume Next
    Set rs = Nothing
    Set con = Nothing
    Exit Function

HandleButtonClick_Err:
    ' If the action was cancelled by the user for
    ' some reason, don't display an error message.
    ' Instead, resume on the next line.
    If (Err = conErrDoCmdCancelled) Then
        Resume Next
    Else
        MsgBox "There was an error executing the command.", vbCritical
        Resume HandleButtonClick_Exit
    End If

End Function

Private Sub cmdLogout_Click()
If gcfHandleErrors Then On Error GoTo Err_cmdLogout_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    'goto main menu
    Me.Filter = "[ItemNumber] = 0 AND [SwitchboardID]=1 " '& rs![Argument]
    DoCmd.Close
    If gcfHandleErrors Then On Error Resume Next
    PunchTheClock glnCashierID, 1
    If gcfHandleErrors Then On Error GoTo Err_cmdLogout_Click

    glnCashierID = 0
    stDocName = "Login"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    If glnCashierID > 0 Then
        Forms("Login").Visible = False
    End If
    If Get_User_Name = True Then
        Shell "c:\kmds\kmdslogout.bat", vbNormalFocus
        Exit Sub
        'DoCmd.Quit

    End If

Exit_cmdLogout_Click:
    Exit Sub

Err_cmdLogout_Click:
    MsgBox Err.Description
    Resume Exit_cmdLogout_Click

End Sub
