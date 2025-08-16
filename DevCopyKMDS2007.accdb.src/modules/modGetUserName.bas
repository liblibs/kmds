Attribute VB_Name = "modGetUserName"
    ' Makes sure all variables are dimensioned in each subroutine.
Option Explicit
    'copied from microsoft support page...

     ' Access the GetUserNameA function in advapi32.dll and
     ' call the function GetUserName.
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
     (ByVal lpBuffer As String, nSize As Long) As Long

     ' Main routine to Dimension variables, retrieve user name
     ' and display answer.

Function Get_User_Name() As Boolean

'check to see if UserName is the same as a cashier number... if so do a autologin
' Dimension variables
    Dim lpBuff As String * 25
    Dim ret As Long, UserName As String
    Dim strUserName As String
    Dim strPassString As String


'Dim comQDF As New Command
Dim comQDF As Object
Set comQDF = CreateObject("ADODB.COMMAND")

'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

    If gcfHandleErrors Then On Error GoTo ErrLog
    'explicit to avoid confusion!
    Get_User_Name = False
    ' Get the user name minus any trailing spaces found in the name.
    ret = GetUserName(lpBuff, 25)
    gstrWindowsUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

    If IsNumeric(gstrWindowsUserName) Then
        With comQDF
            .ActiveConnection = CurrentProject.Connection
            strPassString = "Select * from Cashier where ID = " & gstrWindowsUserName
            .CommandText = strPassString
            .CommandType = adCmdText
            'ADO Numeric Datatypes are very particular
            Set rstemp = .Execute
            If Not rstemp.EOF Then
                rstemp.MoveFirst
                gstrWindowsUserID = rstemp!ID
                gstrWindowsUserPassword = Nz(rstemp!Password)
                Get_User_Name = True
                'jumpout now
                Exit Function
            End If
        End With
        rstemp.Close
        Set rstemp = Nothing
        Set comQDF = Nothing
    Else
       With comQDF
            .ActiveConnection = CurrentProject.Connection
            strPassString = "Select * from Cashier where name = '" & gstrWindowsUserName & "'"
            .CommandText = strPassString
            .CommandType = adCmdText
            'ADO Numeric Datatypes are very particular
            Set rstemp = .Execute
            If Not rstemp.EOF Then
                rstemp.MoveFirst
                gstrWindowsUserID = rstemp!ID
                gstrWindowsUserPassword = Nz(rstemp!Password)
                Get_User_Name = True
                'jumpout now
                Exit Function
            End If
        End With
        rstemp.Close
        Set rstemp = Nothing
        Set comQDF = Nothing
    End If
ExitFunc:
    Exit Function

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description & vbCrLf & strPassString, vbInformation, "Get_User_Name"  ' Display the User Name
    Resume ExitFunc
End Function

Function KMDSEncrypt(RawString As String) As String
Dim strInput As String
Dim strOutput As String
Dim nI As Integer
Dim nChr As Integer
Dim strCharRaw As String
Dim strCharEnc As String
Dim bEven As Boolean
Dim strCharsAllowed As String

If gcfHandleErrors Then On Error GoTo ErrLog
    strCharsAllowed = "abcdefghijklmnopqrstuvwxyz1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    strInput = Trim(RawString)
    For nI = 1 To Len(strInput)
        If (nI \ 2) * 2 = nI Then
            bEven = True
        Else
            bEven = False
        End If
        strCharRaw = Mid(strInput, nI, 1)
        If InStr(strCharsAllowed, strCharRaw) = 0 Then
            MsgBox strCharRaw & vbCrLf & " character not allowed...", vbCritical, "Only letters or numbers!"
            GoTo ErrLog
        End If
        nChr = Asc(strCharRaw)
        If bEven Then
            strCharEnc = Chr(nChr + nI + 50)
            'If not gcfHandleErrors then Debug.Print nChr + ni + 50
        Else
            strCharEnc = Chr(nChr + nI - 30)
            'If not gcfHandleErrors then Debug.Print nChr + ni - 30
        End If
        strOutput = strOutput & strCharEnc

    Next nI
    KMDSEncrypt = strOutput
ExitFunc:
    Exit Function

ErrLog:
    KMDSEncrypt = "NoParse"
    Resume ExitFunc
End Function


Function KMDSDecrypt(EncString As String) As String
Dim strInput As String
Dim strOutput As String
Dim nI As Integer
Dim nChr As Integer
Dim strCharRaw As String
Dim strCharEnc As String
Dim bEven As Boolean

If gcfHandleErrors Then On Error GoTo ErrLog
    strInput = Trim(EncString)
    For nI = 1 To Len(strInput)
        If (nI \ 2) * 2 = nI Then
            bEven = True
        Else
            bEven = False
        End If
        strCharRaw = Mid(strInput, nI, 1)
        nChr = Asc(strCharRaw)
        If bEven Then
            strCharEnc = Chr(nChr - nI - 50)
        Else
            strCharEnc = Chr(nChr - nI + 30)
        End If
        strOutput = strOutput & strCharEnc
    Next nI
    KMDSDecrypt = strOutput

ExitFunc:
    Exit Function

ErrLog:
    KMDSDecrypt = "NoParse"
    Resume ExitFunc
End Function

Function KMDSPasswordCheck(strPassword As String) As Boolean

'criteria for passing
' at least 7 chars long
' at most 10 chars long
' contains at least 1 alpha and 1 number
' contains only alpha and numbers
' has not been used in last 4 passwords


Dim nI As Integer
Dim nChr As Integer

Dim strCharRaw As String
Dim strCharEnc As String
Dim strCharacters As String
Dim strNumbers As String
Dim strSQL As String
Dim strInput As String
Dim strOutput As String

Dim bEven As Boolean
Dim bStringPresent As Boolean
Dim bNumberPresent As Boolean
Dim strPasswordEncrypt As String


'Dim rstemp As New ADODB.Recordset
Dim rstemp As Object
Set rstemp = CreateObject("ADODB.Recordset")

If gcfHandleErrors Then On Error GoTo ErrLog

    If Not gbIntegratedCreditCards Then 'test for complexity
        KMDSPasswordCheck = True
        GoTo ExitFunc
    End If

    KMDSPasswordCheck = False

    bStringPresent = False
    bNumberPresent = False

    If Len(strPassword) < 7 Then
        MsgBox "Passwords must be AT LEAST 7 characters long!", vbCritical, "KMDS Password Check FOR Credit Card processing!"
        GoTo ExitFunc
    End If

    If Len(strPassword) > 10 Then
        MsgBox "Passwords must be LESS THAN 11 characters long!", vbCritical, "KMDS Password Check FOR Credit Card processing!"
        GoTo ExitFunc
    End If

    strCharacters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    strNumbers = "1234567890"
    strInput = Trim(strPassword)
    For nI = 1 To Len(strInput)
        strCharRaw = Mid(strInput, nI, 1)
        If InStr(strCharacters, strCharRaw) = 0 Then
            If InStr(strNumbers, strCharRaw) = 0 Then
                MsgBox strCharRaw & vbCrLf & "Passwords must only contain alphabet or number characters!", vbCritical, "KMDS Password Check FOR Credit Card processing!"
                GoTo ExitFunc
            End If
        End If

        If InStr(strCharacters, strCharRaw) <> 0 Then
            bStringPresent = True
        End If
        If InStr(strNumbers, strCharRaw) <> 0 Then
            bNumberPresent = True
        End If
    Next nI
    If bNumberPresent And bStringPresent Then
        'all other cases return false
        'keep on going if both are true
    Else
        MsgBox "Passwords must be contain at least 1 letter and 1 number!", vbCritical, "KMDS Password Check FOR Credit Card processing!"
        GoTo ExitFunc
    End If

    'if code has gotten this far then it is a valid (complexity wise) password..
    'now check to see if it has been used before (last 4 passwords are used
    '

    If gcfHandleErrors Then On Error GoTo ErrLog

    'MsgBox "Open Cashsys"

    strSQL = "SELECT ID, Password, PasswordHistory1, "
    strSQL = strSQL & "PasswordHistory2, PasswordHistory3, PasswordHistory4, PasswordDate "
    strSQL = strSQL & "From Cashier "
    strSQL = strSQL & "WHERE (ID)=" & CStr(glnCashierIDForPasswordForm)
    With rstemp
        .Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rstemp.BOF And rstemp.EOF Then
            MsgBox "There is no Cashier setup for " & glnCashierIDForPasswordForm
            GoTo ExitFunc
        End If

        strPasswordEncrypt = KMDSEncrypt(strPassword)
        If StrComp(strPasswordEncrypt, !Password, 0) = 0 Then
            MsgBox "Password is the same as one of the last 4!"
            GoTo ExitFunc
        End If
        If StrComp(strPasswordEncrypt, !PasswordHistory1, 0) = 0 Then
            MsgBox "Password is the same as one of the last 4!", vbCritical, "Password1"
            GoTo ExitFunc
        End If
        If StrComp(strPasswordEncrypt, !PasswordHistory2, 0) = 0 Then
            MsgBox "Password is the same as one of the last 4!", vbCritical, "Password2"
            GoTo ExitFunc
        End If
        If StrComp(strPasswordEncrypt, !PasswordHistory3, 0) = 0 Then
            MsgBox "Password is the same as one of the last 4!", vbCritical, "Password3"
            GoTo ExitFunc
        End If
        If StrComp(strPasswordEncrypt, !PasswordHistory4, 0) = 0 Then
            MsgBox "Password is the same as one of the last 4!", vbCritical, "Password4"
            GoTo ExitFunc
        End If
        !PasswordHistory4 = !PasswordHistory3
        !PasswordHistory3 = !PasswordHistory2
        !PasswordHistory2 = !PasswordHistory1
        !PasswordHistory1 = !Password
        !Password = strPasswordEncrypt
        !PasswordDate = Now
        .Update
        .Close
    End With
    Set rstemp = Nothing
    KMDSPasswordCheck = True

ExitFunc:
    Exit Function

ErrLog:
    MsgBox Err.Number & " " & Err.Description, , "KMDSPasswordCheck"
    Resume ExitFunc
End Function
