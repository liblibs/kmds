Attribute VB_Name = "UPCModule"
Option Compare Database
Option Explicit

'revised 08.17.01 sk upcaToe routine added
'                     will take UPC-A and convert to UPC-E (10->6)
'
'revised 02.26.2002 ae  Fixed the 8 digit large label print in Print Labels
'
'revised 04.25.02  sk  added Function UPCECheckdigit(UPCE As String)
'                       This function will return the check digit for UPCE item..
'                       Note: the item must be 6 digits long OR 7 digits with 1st char = "0"


Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3

Global glrsUPCMaster As Recordset
Global glrsPOSLookup As Recordset
Global glrsInvoiceLines  As Recordset
Global glrsMixMatch  As Recordset
Global glrsTaxRate As Recordset
Global glrsDepartment As Recordset
Global glrsCashRegisterSystem As Recordset
Global glrsInvoiceHeader As Recordset
Global glrsSpecials As Recordset
Global glrsPackageDeal As Recordset

Type CashRegisterLineItem
    ItemID As Long
    Description As String
    DepartmentTypeID As Byte
    DepartmentID As Integer
    Quantity As Double
    ExtendedPrice As Currency
    Taxable As Boolean
    TaxRate As Double
    Pricelevel As Byte
End Type                ' female, 1 for male).


Global glcrlInvoiceLineItem As CashRegisterLineItem

   'Declare API functions
   '32bit function for Windows NT and Windows 95


Declare Function sndPlaySound32 Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Declare Function ShowWindow Lib "User32" (ByVal Hwnd As Long, ByVal nCmdShow As Integer) As Integer

Function DFRUNUp()

            'Const lpClassName = "SciCalc"
            Const lpCaption = "DFRUN.EXE"

            'This demonstrates three different ways to call FindWindow:
               '1. The ClassName only.
               '2. The Caption only.
               '3. Both the ClassName and the Caption
               'NOTE: When using Access 2.0, the value 0& must be used to
               '      pass a null string to the FindWindow function rather
               '      than the constant VBNullString.  For Example:
               '      FindWindow(lpClassName, 0&)

            'MsgBox "Calculator Handle = " & FindWindow(lpClassName, _
                  vbNullString)
                  If FindWindow(vbNullString, lpCaption) > 0 Then
                    MsgBox "dfrun.exe is running! HANDLE = " & FindWindow(vbNullString, lpCaption)
                End If
            'MsgBox "Calculator Handle = " & FindWindow(lpClassName, _
                  lpCaption)

            'This function could return the handle of a window.
            'CalculatorUp = FindWindow(lpClassName, 0&)
End Function



Public Function LoadPOSTables()
    Screen.MousePointer = 11
    Set glrsUPCMaster = CurrentDb().OpenRecordset("UPCMaster")
    'Set glrsPOSLookup = CurrentDb().OpenRecordset("SELECT * FROM POSLookup WHERE ActiveAtCashRegister = yes")
    Set glrsPOSLookup = CurrentDb().OpenRecordset("POSLookup")
    Set glrsInvoiceLines = CurrentDb().OpenRecordset("InvoiceLines")
    Set glrsInvoiceHeader = CurrentDb().OpenRecordset("InvoiceHeader")
    Set glrsMixMatch = CurrentDb().OpenRecordset("MixMatch")
    Set glrsTaxRate = CurrentDb().OpenRecordset("TaxRate")
    Set glrsDepartment = CurrentDb().OpenRecordset("SalesDepartment")
    Set glrsCashRegisterSystem = CurrentDb().OpenRecordset("Select * from CashRegisterSystem Where ID =" & 1)
    Set glrsSpecials = CurrentDb().OpenRecordset("Specials")
    Set glrsPackageDeal = CurrentDb().OpenRecordset("PackageDeal")
    Screen.MousePointer = 0
End Function

Public Function ParseUPC(ScanCode As String) As String
    ParseUPC = -1
    Select Case Left(ScanCode, 1)
    Case "A"
        If Len(ScanCode) = 11 Then
            ParseUPC = Mid(ScanCode, 2, 10)
        Else
            ParseUPC = Mid(ScanCode, 3, 10)
        End If
    Case "E"
        If Len(ScanCode) = 7 Then
            ParseUPC = Mid(ScanCode, 2, 6)
        Else
            ParseUPC = Mid(ScanCode, 3, 6)
        End If
    End Select
End Function

Public Function LookUPC(UPCCode As String) As Double
    Dim strCriteria As String
    'glrsUPCMaster.Index = "PrimaryKey"
    If gcfHandleErrors Then On Error GoTo Err_LookUPC

    With glrsUPCMaster
        .Index = "ID"
        .Seek "=", UPCCode '& " '" 'strcriteria)
        If .NoMatch Then
            LookUPC = UPCCode
        Else
            LookUPC = !MasterID
        End If
    End With

    With glrsPOSLookup
        .Index = "PrimaryKey"
        .Seek "=", LookUPC '& " '" 'strcriteria)
        If .NoMatch Then
            LookUPC = -1
        Else
            If glrsPOSLookup!ActiveAtCashRegister = True Then
                LookUPC = !MasterID
            Else
                LookUPC = -1
            End If
        End If
    End With

Exit_LookUPC:
    Exit Function

Err_LookUPC:
    If Err.Number = 3421 Then
        LookUPC = -1
        Resume Exit_LookUPC
    Else
        MsgBox "Error #" & Err.Number & " " & Err.Description, vbExclamation, "Err_LookUPC"
        Resume Exit_LookUPC
    End If
End Function

Public Function GetTaxRate(DepartmentID As Integer) As Double
    With glrsDepartment
        .Index = "ID"
        .Seek "=", DepartmentID '& " '" 'strcriteria)
        If .NoMatch Then
            GetTaxRate = -1
        Else
            glrsTaxRate.Index = "ID"
            glrsTaxRate.Seek "=", glrsDepartment!TaxRateID '& " '" 'strcriteria)
            If glrsTaxRate.NoMatch Then
                GetTaxRate = 0
            Else
                GetTaxRate = glrsTaxRate!Rate
            End If
        End If
    End With
End Function

Public Function LookSpecial(MasterID As Double) As Currency
    Dim strCriteria As String
    'glrsUPCMaster.Index = "PrimaryKey"
    If gcfHandleErrors Then On Error GoTo Err_LookSpecial
    With glrsSpecials
        .Index = "PrimaryKey"
        .Seek "=", MasterID '& " '" 'strcriteria)
        If .NoMatch Then
            LookSpecial = -1
        Else
            If !StartDate < Now And !EndDate > Now Then
                LookSpecial = !Price
            Else
                LookSpecial = -1
            End If
        End If

    End With

Exit_LookSpecial:
    Exit Function

Err_LookSpecial:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbExclamation, "Err_LookSpecial"
    Resume Exit_LookSpecial
End Function


'Public Sub SetOptions(Confirm As Boolean)
'        Application.SetOption "Confirm Record changes", Confirm
'        Application.SetOption "Confirm Document deletions", Confirm
'        Application.SetOption "Confirm Action queries", Confirm
'End Sub
Public Function CheckDigit11(ElevenOutOfTwelve As String) As String
'   MOVE 0 TO CHECK_DIGIT
'   move code to upc_code_str
'   MOVE 0 TO ODD_NUMS
'   MOVE 0 TO EVEN_NUMS
'   FOR COUNTER FROM 1 TO 6
'      MID UPC_CODE_STR TO WORKSTRING 1 ((COUNTER * 2) - 1)
'      CALC (ODD_NUMS + WORKSTRING) TO ODD_NUMS
'      IF COUNTER LT 6 BEGIN
'         MID UPC_CODE_STR TO WORKSTRING 1 (COUNTER * 2)
'         CALC (EVEN_NUMS + WORKSTRING) TO EVEN_NUMS
'         End
'      Loop
'   CALC ((ODD_NUMS * 3) + EVEN_NUMS) TO CHECK_DIGIT
'   RIGHT CHECK_DIGIT TO WORKSTRING 1
'   IF WORKSTRING EQ 0 MOVE 0 TO CHECK_DIGIT
'   ELSE CALC (10 - WORKSTRING) TO CHECK_DIGIT
'   append code check_digit

Dim nI As Integer
Dim strChar As String
Dim nOddNums As Integer
Dim nEvenNums As Integer
Dim nCheckDigit As Integer
If gcfHandleErrors Then On Error GoTo Err_CheckDigit11
CheckDigit11 = ElevenOutOfTwelve

If Len(ElevenOutOfTwelve) = 11 Then
    For nI = 1 To 6
        strChar = Mid(ElevenOutOfTwelve, (nI * 2) - 1, 1)
        nOddNums = nOddNums + CInt(strChar)
        If nI < 6 Then
            strChar = Mid(ElevenOutOfTwelve, (nI * 2), 1)
            nEvenNums = nEvenNums + CInt(strChar)
        End If
    Next nI
    nCheckDigit = (nOddNums * 3) + nEvenNums
    strChar = Right(CStr(nCheckDigit), 1)

    If CInt(strChar) = 0 Then
        nCheckDigit = 0
    Else
        nCheckDigit = 10 - CInt(strChar)
    End If
    CheckDigit11 = ElevenOutOfTwelve & CStr(nCheckDigit)
End If
Exit_CheckDigit11:
    Exit Function
Err_CheckDigit11:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbExclamation, "CheckDigit11 " & ElevenOutOfTwelve
End Function

Public Function ExpandUPCE(UPCE As String) As String
'Page scrn1
'
'string upc_code_str workstring tmp_str upc_a
'integer code_length sys_digit chk_digit
'
'
'top:
'accept scrn1.1
'move scrn1.1 to upc_code_str
'length upc_code_str to code_length
'move "" to scrn1.2
'GoSub MAKE_UPC_A
'move upc_code_str to scrn1.2
'GoTo top
'
'keyproc Key.escape
'System
'Return
'
'
'
'MAKE_UPC_A:
'   IF CODE_LENGTH EQ 8 BEGIN
'      LEFT UPC_CODE_STR TO SYS_DIGIT 1
'      RIGHT UPC_CODE_STR TO CHK_DIGIT 1
'      MID UPC_CODE_STR TO UPC_CODE_STR 6 2
'      End
'   RIGHT UPC_CODE_STR TO TMP_STR 1
'   IF TMP_STR LE 2 BEGIN
'      LEFT UPC_CODE_STR TO UPC_A 2
'      MID UPC_CODE_STR TO UPC_CODE_STR 3 3
'      IF TMP_STR EQ 0 APPEND UPC_A '00000' UPC_CODE_STR
'      IF TMP_STR EQ 1 APPEND UPC_A '10000' UPC_CODE_STR
'      IF TMP_STR EQ 2 APPEND UPC_A '20000' UPC_CODE_STR
'      MOVE UPC_A TO UPC_CODE_STR
'      End
'   ELSE IF TMP_STR EQ 3 BEGIN
'      LEFT UPC_CODE_STR TO UPC_A 3
'      MID UPC_CODE_STR TO UPC_CODE_STR 2 4
'      APPEND UPC_A '00000' UPC_CODE_STR
'      MOVE UPC_A TO UPC_CODE_STR
'      End
'   ELSE IF TMP_STR EQ 4 BEGIN
'      LEFT UPC_CODE_STR TO UPC_A 4
'      MID UPC_CODE_STR TO UPC_CODE_STR 1 5
'      APPEND UPC_A '00000' UPC_CODE_STR
'      MOVE UPC_A TO UPC_CODE_STR
'      End
'   ELSE IF TMP_STR GT 4 BEGIN
'      LEFT UPC_CODE_STR TO UPC_A 5
'      MID UPC_CODE_STR TO UPC_CODE_STR 1 6
'      APPEND UPC_A '0000' UPC_CODE_STR
'      MOVE UPC_A TO UPC_CODE_STR
'      End
'   IF CODE_LENGTH EQ 8 BEGIN
'      MOVE "" TO WORKSTRING
'      APPEND WORKSTRING SYS_DIGIT UPC_CODE_STR CHK_DIGIT
'      MOVE WORKSTRING TO UPC_CODE_STR
'      End
'   Return
'

Dim upc_code_str As String
Dim workstring As String
Dim tmp_str As String
Dim upc_a As String

Dim code_length As Integer
Dim sys_digit As Integer
Dim chk_digit As Integer

If Len(UPCE) <> 6 And Len(UPCE) <> 8 Then
    ExpandUPCE = UPCE
    Exit Function
End If

upc_code_str = UPCE
code_length = Len(upc_code_str)

If code_length = 8 Then
    sys_digit = Left(upc_code_str, 1)
    chk_digit = Right(upc_code_str, 1)
    upc_code_str = Mid(upc_code_str, 2, 6)
End If

tmp_str = Right(upc_code_str, 1)
If tmp_str <= 2 Then
    upc_a = Left(upc_code_str, 2)
    upc_code_str = Mid(upc_code_str, 3, 3)
    If tmp_str = 0 Then upc_a = upc_a & "00000" & upc_code_str
    If tmp_str = 1 Then upc_a = upc_a & "10000" & upc_code_str
    If tmp_str = 2 Then upc_a = upc_a & "20000" & upc_code_str
    upc_code_str = upc_a
Else
    If tmp_str = 3 Then
        upc_a = Left(upc_code_str, 3)
        upc_code_str = Mid(upc_code_str, 4, 2)
        upc_a = upc_a & "00000" & upc_code_str
        upc_code_str = upc_a
    Else
        If tmp_str = 4 Then
            upc_a = Left(upc_code_str, 4)
            upc_code_str = Mid(upc_code_str, 5, 1)
            upc_a = upc_a & "00000" & upc_code_str
            upc_code_str = upc_a
        Else
            If tmp_str > 4 Then
                upc_a = Left(upc_code_str, 5)
                upc_code_str = Mid(upc_code_str, 6, 1)
                upc_a = upc_a & "0000" & upc_code_str
                upc_code_str = upc_a
            End If
        End If
    End If
End If


If code_length = 8 Then
    workstring = sys_digit & upc_code_str & chk_digit
    upc_code_str = workstring
End If

ExpandUPCE = upc_code_str
If Len(ExpandUPCE) = 10 Then
    ExpandUPCE = CheckDigit11("0" & ExpandUPCE)
End If


End Function


Public Function BuildUPCAGlyph(UPCA As String) As String
Dim nI As Integer
Dim strTemp As String

strTemp = Chr(Left(UPCA, 1) + 80)
For nI = 1 To 5
    strTemp = strTemp & Chr(Mid(UPCA, nI + 1, 1) + 48)
Next nI
strTemp = strTemp & Chr(112)
For nI = 6 To 10
    strTemp = strTemp & Chr(Mid(UPCA, nI + 1, 1) + 64)
Next nI
strTemp = strTemp & Chr(Right(UPCA, 1) + 96)

BuildUPCAGlyph = strTemp
End Function

Public Function CheckDigit12(TwelveOutOfThirteen As String) As String
'   MOVE 0 TO CHECK_DIGIT
'   move code to upc_code_str
'   MOVE 0 TO ODD_NUMS
'   MOVE 0 TO EVEN_NUMS
'   FOR COUNTER FROM 1 TO 6
'      MID UPC_CODE_STR TO WORKSTRING 1 ((COUNTER * 2) - 1)
'      CALC (ODD_NUMS + WORKSTRING) TO ODD_NUMS
'      IF COUNTER LT 6 BEGIN
'         MID UPC_CODE_STR TO WORKSTRING 1 (COUNTER * 2)
'         CALC (EVEN_NUMS + WORKSTRING) TO EVEN_NUMS
'         End
'      Loop
'   CALC ((ODD_NUMS * 3) + EVEN_NUMS) TO CHECK_DIGIT
'   RIGHT CHECK_DIGIT TO WORKSTRING 1
'   IF WORKSTRING EQ 0 MOVE 0 TO CHECK_DIGIT
'   ELSE CALC (10 - WORKSTRING) TO CHECK_DIGIT
'   append code check_digit

Dim nI As Integer
Dim strChar As String
Dim nOddNums As Integer
Dim nEvenNums As Integer
Dim nCheckDigit As Integer
If gcfHandleErrors Then On Error GoTo Err_CheckDigit12
CheckDigit12 = TwelveOutOfThirteen

If Len(TwelveOutOfThirteen) = 12 Then
    For nI = 1 To 6
        strChar = Mid(TwelveOutOfThirteen, (nI * 2), 1)
        nOddNums = nOddNums + CInt(strChar)
        strChar = Mid(TwelveOutOfThirteen, (nI * 2) - 1, 1)
        nEvenNums = nEvenNums + CInt(strChar)

    Next nI
    nCheckDigit = (nOddNums * 3) + nEvenNums
    strChar = Right(CStr(nCheckDigit), 1)

    If CInt(strChar) = 0 Then
        nCheckDigit = 0
    Else
        nCheckDigit = 10 - CInt(strChar)
    End If
    CheckDigit12 = TwelveOutOfThirteen & CStr(nCheckDigit)
End If
Exit_CheckDigit12:
    Exit Function
Err_CheckDigit12:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbExclamation, "CheckDigit12 " & TwelveOutOfThirteen
End Function

Public Function BuildEANGlyph(EAN As String) As String

'WARNING:
'
'This font contains only the glyphs to BUILD UPC-A barcodes,
' which must be assembled manually or through software to
' obtain the final mechanicaly readable product.
'In order for a barcode reader to properly recognize the barcode,
' the following instructions must be followed exactly.
'
'EAN Codes:
'
'The EAN Codes encode either 13 or 8 characters.
'The 13 character version is a superset of the UPC A code.
'The 8 character version is for printing on smaller packages.
'
'EAN 13:
'
'EAN 13 is used to encode thirteen characters.
' The first two characters are the flag characters that identify
' the country of origin,
' the next ten characters are the data characters,
' and the last character is the check character.
'
'NOTE
'
'The UPC-A symbol is a subset of the EAN 13 symbols.
' A UPC-A symbol is an EAN 13 symbol with the first flag character set to zero.
'
'The physical structure of the EAN-13 is as follows:
'
'    Left hand guard bars, encoded 101
'    Second flag character, encoded as in the following tables
'    First five data characters, encoded as in the following tables
'    Center guard pattern, encoded 01010
'    Last five data characters, encoded as right hand characters
'    Check character, encoded as a right hand character
'    Right hand guard bars, encoded 101
'
'    Nominal module bar/space width = 0.013 in (0.33 mm)
'    Worst case module width = 0.009 in. (0.23 mm)
'
'The thirteenth digit (First Flag Character) is encoded in the
' pattern of parity in the second flag character
' and the first five data characters.
'The left hand A and the left hand B character sets are used.
'
'EAN Character Set:
'
'Data  Left Hand A  Left Hand B  Right Hand
'
'  0     0001101      0100111      1110010
'  1     0011001      0110011      1100110
'  2     0010011      0011011      1101100
'  3     0111101      0100001      1000010
'  4     0100011      0011101      1011100
'  5     0110001      0111001      1001110
'  6     0101111      0000101      1010000
'  7     0111011      0010001      1000100
'  8     0110111      0001001      1001000
'  9     0001011      0010111      1110100
'
'First Flag Character Encodation for EAN 13
'
'First Flag  Second Flag Data Characters
'Character   Character   1  2  3  4  5
'
'    0          A        A  A  A  A  A
'    1          A        A  B  A  B  B
'    2          A        A  B  B  A  B
'    3          A        A  B  B  B  A
'    4          A        B  A  A  B  B
'    5          A        B  B  A  A  B
'    6          A        B  B  B  A  A
'    7          A        B  A  B  A  B
'    8          A        B  A  B  B  A
'    9          A        B  B  A  B  A
'
'Check characters:
'
'The check character for the UPC/EAN codes is generated by the following method:
'
'1 - Designate the rightmost character odd.
'2 - Sum all of the characters in the odd positions and multiply the result by
'    three. (Remember to include the Number System Character)
'3 - Sum all of the characters in the even positions
'4 - Add the odd and even totals from steps two and three.
'5 - Determine the smallest number that when added to the result from step
'    four, will result in a multiple of 10. This is the check character.
'
'EAN 13 Example
'
'Flag characters = 20
'Message characters = 0123456789
'Position = EOEOEOEOEOEO
'Sum of odd positions =  0 + 1 + 3 + 5 + 7 + 9 = 25
'Odd positions X 3 =     25 x 3 = 75
'Sum of even positions = 2 + 0 + 2 + 4 + 6 + 8 = 22
'Sum of even and odd =   75 + 22 = 97
'Check character = 3
'Tag to be encoded =     2001234567893
'
'Left Hand A barcode glyphs 0 to 9 are at 48  (0x30) to 57 (0x39)
'Left Hand B barcode glyphs 0 to 9 are at 64  (0x40) to 73 (0x49).
'Right Hand barcode  glyphs 0 to 9 are at 80  (0x50) to 89 (0x59).
'First Flag barcode  glyphs 0 to 9 are at 33  (0x21) to 42 (0x2A).
'The Secong Flag Characters 0 to 9 are at 96  (0x60) to 105 (0x69).
'Check Characters'                 are at 112 (0x70) to 121 (0x79).
'Center Guard' glyph                is at 124 (0x7C).
'First Flag  Second Flag Data Characters
'Character   Character   1  2  3  4  5
'
'    0          A        A  A  A  A  A
'    1          A        A  B  A  B  B
'    2          A        A  B  B  A  B
'    3          A        A  B  B  B  A
'    4          A        B  A  A  B  B
'    5          A        B  B  A  A  B
'    6          A        B  B  B  A  A
'    7          A        B  A  B  A  B
'    8          A        B  A  B  B  A
'    9          A        B  B  A  B  A

'Second Flag Characters and Check Characters
'  also include the Left and Right Hand Guard Bars.

Dim nI As Integer
Dim strTemp As String
Dim strFirstFlag As String
Dim strSecondFlag As String
Dim strLeftChar As String
Dim nLeftChar As Integer
strFirstFlag = Mid(EAN, 1, 1)
strSecondFlag = Mid(EAN, 2, 1)

strTemp = Chr(Mid(EAN, 1, 1) + 33)
strTemp = strTemp & Chr(Mid(EAN, 2, 1) + 96)


For nI = 1 To 5
    nLeftChar = Mid(EAN, (nI + 2), 1)

    Select Case strFirstFlag
'    0          A        A  A  A  A  A
'    1          A        A  B  A  B  B
'    2          A        A  B  B  A  B
'    3          A        A  B  B  B  A
'    4          A        B  A  A  B  B
'    5          A        B  B  A  A  B
'    6          A        B  B  B  A  A
'    7          A        B  A  B  A  B
'    8          A        B  A  B  B  A
'    9          A        B  B  A  B  A
    Case "0"
        strLeftChar = Chr(nLeftChar + 48)
    Case "1"
        Select Case nI
        Case 1
            strLeftChar = Chr(nLeftChar + 48)
        Case 2
            strLeftChar = Chr(nLeftChar + 64)
        Case 3
            strLeftChar = Chr(nLeftChar + 48)
        Case 4
            strLeftChar = Chr(nLeftChar + 64)
        Case 5
            strLeftChar = Chr(nLeftChar + 64)
        End Select
    Case "2"
        Select Case nI
        Case 1
            strLeftChar = Chr(nLeftChar + 48)
        Case 2
            strLeftChar = Chr(nLeftChar + 64)
        Case 3
            strLeftChar = Chr(nLeftChar + 64)
        Case 4
            strLeftChar = Chr(nLeftChar + 48)
        Case 5
            strLeftChar = Chr(nLeftChar + 64)
        End Select
    Case "3"
        Select Case nI
        Case 1
            strLeftChar = Chr(nLeftChar + 48)
        Case 2
            strLeftChar = Chr(nLeftChar + 64)
        Case 3
            strLeftChar = Chr(nLeftChar + 64)
        Case 4
            strLeftChar = Chr(nLeftChar + 64)
        Case 5
            strLeftChar = Chr(nLeftChar + 48)
        End Select
    Case "4"
        Select Case nI
        Case 1
            strLeftChar = Chr(nLeftChar + 64)
        Case 2
            strLeftChar = Chr(nLeftChar + 48)
        Case 3
            strLeftChar = Chr(nLeftChar + 48)
        Case 4
            strLeftChar = Chr(nLeftChar + 64)
        Case 5
            strLeftChar = Chr(nLeftChar + 64)
        End Select
    Case "5"
        Select Case nI
        Case 1
            strLeftChar = Chr(nLeftChar + 64)
        Case 2
            strLeftChar = Chr(nLeftChar + 64)
        Case 3
            strLeftChar = Chr(nLeftChar + 48)
        Case 4
            strLeftChar = Chr(nLeftChar + 48)
        Case 5
            strLeftChar = Chr(nLeftChar + 64)
        End Select
    Case "6"
        Select Case nI
        Case 1
            strLeftChar = Chr(nLeftChar + 64)
        Case 2
            strLeftChar = Chr(nLeftChar + 64)
        Case 3
            strLeftChar = Chr(nLeftChar + 64)
        Case 4
            strLeftChar = Chr(nLeftChar + 48)
        Case 5
            strLeftChar = Chr(nLeftChar + 48)
        End Select
    Case "7"
        Select Case nI
        Case 1
            strLeftChar = Chr(nLeftChar + 64)
        Case 2
            strLeftChar = Chr(nLeftChar + 48)
        Case 3
            strLeftChar = Chr(nLeftChar + 64)
        Case 4
            strLeftChar = Chr(nLeftChar + 48)
        Case 5
            strLeftChar = Chr(nLeftChar + 64)
        End Select
    Case "8"
        Select Case nI
        Case 1
            strLeftChar = Chr(nLeftChar + 64)
        Case 2
            strLeftChar = Chr(nLeftChar + 48)
        Case 3
            strLeftChar = Chr(nLeftChar + 64)
        Case 4
            strLeftChar = Chr(nLeftChar + 64)
        Case 5
            strLeftChar = Chr(nLeftChar + 48)
        End Select
    Case "9"
        Select Case nI
        Case 1
            strLeftChar = Chr(nLeftChar + 64)
        Case 2
            strLeftChar = Chr(nLeftChar + 64)
        Case 3
            strLeftChar = Chr(nLeftChar + 48)
        Case 4
            strLeftChar = Chr(nLeftChar + 64)
        Case 5
            strLeftChar = Chr(nLeftChar + 48)
        End Select
    End Select
    strTemp = strTemp & strLeftChar
Next nI


strTemp = strTemp & Chr(124)
For nI = 6 To 10
    strTemp = strTemp & Chr(Mid(EAN, nI + 2, 1) + 80)
Next nI
strTemp = strTemp & Chr(Right(EAN, 1) + 112)

BuildEANGlyph = strTemp
End Function

Public Function BuildI2O5Glyph(I2o5String As String) As String
    Dim nI As Integer

    If IsNumeric(I2o5String) And Len(I2o5String) > 0 Then
        If ((Len(I2o5String) \ 2) * 2) < Len(I2o5String) Then
            I2o5String = "0" & I2o5String
        End If
        BuildI2O5Glyph = "("
        For nI = 0 To (Len(I2o5String) / 2 - 1)
            BuildI2O5Glyph = BuildI2O5Glyph & Chr(Mid(I2o5String, (nI * 2) + 1, 2) + Abs(Mid(I2o5String, (nI * 2) + 1, 2) <= 49) * 48 + Abs(Mid(I2o5String, (nI * 2) + 1, 2) >= 50) * 142)
        Next nI
        BuildI2O5Glyph = BuildI2O5Glyph & ")"
    End If
End Function

Public Sub LoadGFTPServer()
    'Look for the gftp server program... if not loaded then load it
    If FindWindow(vbNullString, "G-FTP Server") = 0 Then
        If Shell("C:\gsite\gftps.exe", vbNormalFocus) = 0 Then
            MsgBox "Can not find file named: " & "C:\gsite\gftps.exe"
        Else
            'it takes about 5 seconds to load on my machine
            'so i'll wait for 10
            Sleep 10000
        End If
    End If
End Sub

Function PrintLabels(strStyle As String, dblQuantity As Double, strDescription As String, strPrice As String, strCode As String, dblLength As Double) As String
    Dim strTFirst As String
    Dim strTFourth As String
    Dim strTSecond As String
    Dim strTThird As String

    If dblLength = 12 Then
        strTFirst = Left(strCode, 1)
        strTFourth = Right(strCode, 11)
        strTSecond = Left(strTFourth, 5)
        strTFourth = Right(strTFourth, 6)
        strTThird = Left(strTFourth, 5)
        strTFourth = Right(strTFourth, 1)
    End If

    Select Case strStyle
        Case Is = 118
            strCode = Mid(strCode, 2, 6)
            strDescription = Left(strDescription, 18)
            PrintLabels = "! 0 100 80 " & dblQuantity & vbCrLf
            PrintLabels = PrintLabels & "PITCH 100" & vbCrLf
            PrintLabels = PrintLabels & "WIDTH 240" & vbCrLf
            PrintLabels = PrintLabels & "MULTIPLE 2" & vbCrLf
            PrintLabels = PrintLabels & "STRING 5X7 0 0 " & strDescription & vbCrLf
            PrintLabels = PrintLabels & "STRING 12X16 0 62 " & strPrice & vbCrLf
            PrintLabels = PrintLabels & "BARCODE UPCE 30 50 40 " & strCode & vbCrLf
            PrintLabels = PrintLabels & "INDEX" & vbCrLf
            PrintLabels = PrintLabels & "END"
        Case Is = 1112
            PrintLabels = "! 0 100 180 " & dblQuantity & vbCrLf
            PrintLabels = PrintLabels & "PITCH 200" & vbCrLf
            PrintLabels = PrintLabels & "WIDTH 240" & vbCrLf
            PrintLabels = PrintLabels & "MULTIPLE 2" & vbCrLf
            PrintLabels = PrintLabels & "STRING 5X7(2,3) 15 0 " & strDescription & vbCrLf
            PrintLabels = PrintLabels & "STRING 18X23 25 135 " & strPrice & vbCrLf
            PrintLabels = PrintLabels & "BT 12X16" & vbCrLf
            PrintLabels = PrintLabels & "BARCODE UPCA-(2,3) 15 95 70 " & strCode & vbCrLf
            PrintLabels = PrintLabels & "STRING 5X7(2,3) 10 99 " & strTFirst & " " & strTSecond & " " & strTThird & " " & strTFourth & vbCrLf
            PrintLabels = PrintLabels & "INDEX" & vbCrLf
            PrintLabels = PrintLabels & "END"
        Case Is = 128
            PrintLabels = Mid(strCode, 2, 6) & vbCrLf
            PrintLabels = PrintLabels & "! 0 100 95 " & dblQuantity & vbCrLf
            PrintLabels = PrintLabels & "PITCH 100" & vbCrLf
            PrintLabels = PrintLabels & "WIDTH 224" & vbCrLf
            PrintLabels = PrintLabels & "STRING 9X12 0 0 " & strDescription & vbCrLf
            'strLine4 = "STRING 5X7 10 80 Category #"
            PrintLabels = PrintLabels & "STRING 12X16 125 60 " & strPrice & vbCrLf
            PrintLabels = PrintLabels & "BARCODE UPCE 20 65 40 " & Mid(strCode, 2, 6) & vbCrLf
            PrintLabels = PrintLabels & "INDEX" & vbCrLf
            PrintLabels = PrintLabels & "END"
        Case Is = 1212
            PrintLabels = "! 0 100 95 " & dblQuantity & vbCrLf
            PrintLabels = PrintLabels & "PITCH 100" & vbCrLf
            PrintLabels = PrintLabels & "WIDTH 224" & vbCrLf
            PrintLabels = PrintLabels & "STRING 9X12 0 0" & strDescription & vbCrLf
            'strLine4 = "STRING 5X7 0 80 Category #"
            PrintLabels = PrintLabels & "STRING 12X16 125 60 " & strPrice & vbCrLf
            PrintLabels = PrintLabels & "BARCODE UPCA- 10 65 40 " & strCode & vbCrLf
            PrintLabels = PrintLabels & "STRING 5X7 10 70 " & strTFirst & " " & strTSecond & " " & strTThird & " " & strTFourth & vbCrLf
            PrintLabels = PrintLabels & "INDEX" & vbCrLf
            PrintLabels = PrintLabels & "END"
        Case Is = 218
            PrintLabels = "! 0 100 80 " & dblQuantity & vbCrLf
            PrintLabels = PrintLabels & "VARIABLE ENCODER 1" & vbCrLf
            PrintLabels = PrintLabels & "PITCH 100" & vbCrLf
            PrintLabels = PrintLabels & "WIDTH 240" & vbCrLf
            PrintLabels = PrintLabels & "MULTIPLE 2" & vbCrLf
            PrintLabels = PrintLabels & "STRING 3X5 5 5 " & strDescription & vbCrLf
            PrintLabels = PrintLabels & "STRING 12X16 5 67 " & strPrice & vbCrLf
            'PrintLabels = PrintLabels & "BT 12X16" & vbCrLf
            PrintLabels = PrintLabels & "BARCODE UPCE 30 55 40 " & strCode & vbCrLf
            PrintLabels = PrintLabels & "INDEX" & vbCrLf
            PrintLabels = PrintLabels & "END"
        Case Is = 2112
            PrintLabels = "! 0 100 125 " & dblQuantity & vbCrLf
            PrintLabels = PrintLabels & "VARIABLE ENCODER 1" & vbCrLf
            PrintLabels = PrintLabels & "PITCH 200" & vbCrLf
            PrintLabels = PrintLabels & "WIDTH 240" & vbCrLf
            PrintLabels = PrintLabels & "MULTIPLE 2" & vbCrLf
            PrintLabels = PrintLabels & "STRING 9X12 5 0 " & strDescription & vbCrLf
            PrintLabels = PrintLabels & "STRING 18X23 5 105 " & strPrice & vbCrLf
            PrintLabels = PrintLabels & "BT 12X16" & vbCrLf
            PrintLabels = PrintLabels & "BARCODE UPCA-(2,3) 5 65 50 " & strCode & vbCrLf
            PrintLabels = PrintLabels & "STRING 12X16 5 76 " & strTFirst & " " & strTSecond & " " & strTThird & " " & strTFourth & vbCrLf
            PrintLabels = PrintLabels & "INDEX" & vbCrLf
            PrintLabels = PrintLabels & "END"
        Case Is = 228
            PrintLabels = "! 0 100 95 " & dblQuantity & vbCrLf
            PrintLabels = PrintLabels & "VARIABLE ENCODER 1" & vbCrLf
            PrintLabels = PrintLabels & "PITCH 100" & vbCrLf
            PrintLabels = PrintLabels & "WIDTH 224" & vbCrLf
            PrintLabels = PrintLabels & "STRING 9X12 0 0 " & strDescription & vbCrLf
            'strLine5 = "STRING 5X7 10 80 Category#" & rsPOSLookup!Category
            PrintLabels = PrintLabels & "STRING 12X16 125 60 " & strPrice & vbCrLf
            PrintLabels = PrintLabels & "BARCODE UPCE 20 65 40 " & strCode & vbCrLf
            PrintLabels = PrintLabels & "INDEX" & vbCrLf
            PrintLabels = PrintLabels & "END"
        Case Is = 2212
            PrintLabels = "! 0 100 95 " & dblQuantity & vbCrLf
            PrintLabels = PrintLabels & "VAIRABLE ENCODER 1" & vbCrLf
            PrintLabels = PrintLabels & "PITCH 100" & vbCrLf
            PrintLabels = PrintLabels & "WIDTH 224" & vbCrLf
            PrintLabels = PrintLabels & "STRING 9X12 0 0 " & strDescription & vbCrLf
            'strLine5 = "STRING 5X7 0 80 Category#" & rsPOSLookup!Category
            PrintLabels = PrintLabels & "STRING 12X16 125 60 " & strPrice & vbCrLf
            PrintLabels = PrintLabels & "BARCODE UPCA+ 10 65 40 " & strCode & vbCrLf
            PrintLabels = PrintLabels & "INDEX" & vbCrLf
            PrintLabels = PrintLabels & "END"
        End Select

End Function


Function UPCAtoE(A2E_INPUT As String) As String
'pass a 10 or 11 or 12 digit number
'if it can be compressed then
'   return compressed (E)
'else
'  return original value
'
Dim A2E_VENDOR As String
Dim A2E_ITEM As String
Dim A2E_OUTPUT As String
Dim A2E_VENDOR_SUPP As String
Dim A2E_VENDOR_3 As String
Dim A2E_CASE As String
Dim CASE_QUALIFIES As String
Dim A2E_ZEROS As String
Dim CheckDigit As String

UPCAtoE = A2E_INPUT

    A2E_VENDOR = Len(A2E_INPUT)

    Select Case A2E_VENDOR
    Case 10
        CheckDigit = Right(CheckDigit11("0" & A2E_INPUT), 1)
    Case 11
        CheckDigit = Left(A2E_INPUT, 1)
        If CheckDigit <> "0" Then Exit Function
        CheckDigit = Right(CheckDigit11(A2E_INPUT), 1)
        A2E_INPUT = Mid(A2E_INPUT, 2, 10)
    Case 12
        CheckDigit = Left(A2E_INPUT, 1)
        If CheckDigit <> "0" Then Exit Function
        CheckDigit = Right(A2E_INPUT, 1)
        A2E_INPUT = Mid(A2E_INPUT, 2, 10)
    End Select

    A2E_VENDOR = Len(A2E_INPUT)


    If A2E_VENDOR = "10" Then
        A2E_VENDOR = Mid(A2E_INPUT, 3, 3)
        A2E_CASE = 0
        If A2E_VENDOR = "000" Then A2E_CASE = 1
        If A2E_VENDOR = "100" Then A2E_CASE = 1
        If A2E_VENDOR = "200" Then A2E_CASE = 1
        If A2E_CASE = 0 Then
            A2E_VENDOR = Mid(A2E_INPUT, 4, 2)
            If A2E_VENDOR = "00" Then A2E_CASE = 2
            If A2E_CASE = 0 Then
                A2E_VENDOR = Mid(A2E_INPUT, 5, 1)
                If CDbl(A2E_VENDOR) = 0 Then
                    A2E_CASE = 3 'EQ 0 MOVE 3 TO A2E_CASE
                Else
                    A2E_CASE = 4 'MOVE 4 TO A2E_CASE
                End If
            End If
        End If
    End If


    A2E_VENDOR = Left(A2E_INPUT, 5)    'TO A2E_VENDOR 5
    A2E_ITEM = CDbl(Right(A2E_INPUT, 5))
    CASE_QUALIFIES = 0

    Select Case A2E_CASE
    Case "1"
        If CDbl(A2E_ITEM) < 1000 Then
            CASE_QUALIFIES = 1
            A2E_VENDOR_SUPP = Left(A2E_VENDOR, 2) 'TO A2E_VENDOR_SUPP 2
            If CDbl(A2E_ITEM) < 100 Then A2E_ZEROS = "0"
            If CDbl(A2E_ITEM) < 10 Then A2E_ZEROS = "00"
            A2E_VENDOR_3 = Mid(A2E_VENDOR, 3, 1)
        End If

    Case "2"
         If CDbl(A2E_ITEM) < 100 Then
            CASE_QUALIFIES = 1
            A2E_VENDOR_SUPP = Left(A2E_VENDOR, 3)
            If CDbl(A2E_ITEM) < 10 Then A2E_ZEROS = "0"
            A2E_VENDOR_3 = 3
        End If

    Case "3"
         If CDbl(A2E_ITEM) < 10 Then
            CASE_QUALIFIES = 1
            A2E_VENDOR_SUPP = Left(A2E_VENDOR, 4)
            A2E_ZEROS = ""
            A2E_VENDOR_3 = 4
        End If

    Case "4"
         If CDbl(A2E_ITEM) >= 5 And CDbl(A2E_ITEM) <= 9 Then
            CASE_QUALIFIES = 1
            A2E_VENDOR_SUPP = A2E_VENDOR
            A2E_ZEROS = ""
            A2E_VENDOR_3 = A2E_ITEM
            A2E_ITEM = ""
        End If

    End Select


    If CASE_QUALIFIES = "1" Then
       UPCAtoE = "0" & A2E_VENDOR_SUPP & A2E_ZEROS & A2E_ITEM & A2E_VENDOR_3 & CheckDigit
    End If

End Function

'NAME A2E_INPUT
'NAME A2E_VENDOR
'NAME A2E_ITEM
'NAME A2E_OUTPUT
'NAME A2E_VENDOR_SUPP
'NAME A2E_VENDOR_3
'NAME A2E_CASE
'NAME CASE_QUALIFIES
'NAME A2E_ZEROS
'REPEAT
'   ACCEPT A2E_INPUT
'   BLANKFORM A2E_VENDOR THRU A2E_ZEROS
'   LENGTH A2E_INPUT TO A2E_VENDOR
'   IF A2E_VENDOR EQ 10 BEGIN
'      MID A2E_INPUT TO A2E_VENDOR 3 3
'      MOVE 0 TO A2E_CASE
'      IF A2E_VENDOR EQ 0   MOVE 1 TO A2E_CASE
'      IF A2E_VENDOR EQ 100 MOVE 1 TO A2E_CASE
'      IF A2E_VENDOR EQ 200 MOVE 1 TO A2E_CASE
'      IF A2E_CASE EQ 0 BEGIN
'         MID A2E_INPUT TO A2E_VENDOR 2 4
'         IF A2E_VENDOR EQ 0 MOVE 2 TO A2E_CASE
'         IF A2E_CASE EQ 0 BEGIN
'            MID A2E_INPUT TO A2E_VENDOR 1 5
'            IF A2E_VENDOR EQ 0 MOVE 3 TO A2E_CASE
'            ELSE MOVE 4 TO A2E_CASE
'            End
'         End
'      LEFT  A2E_INPUT TO A2E_VENDOR 5
'      RIGHT A2E_INPUT TO A2E_ITEM   5
'      MOVE 0 TO CASE_QUALIFIES
'      IF A2E_CASE EQ 1 BEGIN
'         IF A2E_ITEM LT 1000 BEGIN
'            MOVE 1 TO CASE_QUALIFIES
'            LEFT A2E_VENDOR TO A2E_VENDOR_SUPP 2
'            IF A2E_ITEM LT 100 MOVE "0" TO A2E_ZEROS
'            IF A2E_ITEM LT 10  MOVE "00" TO A2E_ZEROS
'            MID A2E_VENDOR TO A2E_VENDOR_3 1 3
'            End
'         End
'      IF A2E_CASE EQ 2 BEGIN
'         IF A2E_ITEM LT 100 BEGIN
'            MOVE 1 TO CASE_QUALIFIES
'            LEFT A2E_VENDOR TO A2E_VENDOR_SUPP 3
'            IF A2E_ITEM LT 10 MOVE "0" TO A2E_ZEROS
'            MOVE 3 TO A2E_VENDOR_3
'            End
'         End
'      IF A2E_CASE EQ 3 BEGIN
'         IF A2E_ITEM LT 10 BEGIN
'            MOVE 1 TO CASE_QUALIFIES
'            LEFT A2E_VENDOR TO A2E_VENDOR_SUPP 4
'            MOVE "" TO A2E_ZEROS
'            MOVE 4 TO A2E_VENDOR_3
'            End
'         End
'      IF A2E_CASE EQ 4 BEGIN
'         IF A2E_ITEM GE 5 IF A2E_ITEM LE 9 BEGIN
'            MOVE 1 TO CASE_QUALIFIES
'            MOVE A2E_VENDOR TO A2E_VENDOR_SUPP
'            MOVE "" TO A2E_ZEROS
'            MOVE A2E_ITEM TO A2E_VENDOR_3
'            End
'         End
'      IF CASE_QUALIFIES EQ 1 BEGIN
'      APPEND A2E_OUTPUT A2E_VENDOR_SUPP A2E_ZEROS A2E_ITEM A2E_VENDOR_3
'      REPEAT
'         REPLACE " " IN  A2E_OUTPUT WITH ""
'         [FOUND] LOOP
'      End
'      End
'   Loop


Public Function UPCECheckdigit(UPCE As String) As String
'This function will return the check digit for UPCE item..
'Note: the item must be 6 digits long OR 7 digits with 1st char = "0"

    If gcfHandleErrors Then On Error GoTo ErrLog

    UPCECheckdigit = UPCE
    Select Case Len(UPCE)
    Case 6
        UPCECheckdigit = "0" & UPCE & Right(ExpandUPCE(UPCE), 1)
    Case 7
        If Left(UPCE, 1) = "0" Then
            UPCECheckdigit = UPCE & Right(ExpandUPCE(Right(UPCE, 6)), 1)
        Else
            'If not gcfHandleErrors then Debug.Print "Not upce!"
        End If
    End Select
ExitFunc:
    'If not gcfHandleErrors then Debug.Print "V:"; UPCECheckdigit
    Exit Function
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "UPCModule:UPCECheckDigit"
    Resume ExitFunc
End Function



Function MaximizeApp()
   Dim Maxit%
   Maxit% = ShowWindow(hWndAccessApp, SW_SHOWMAXIMIZED)
End Function

Function MinimizeApp()
   Dim Minit%
   Minit% = ShowWindow(hWndAccessApp, SW_SHOWMINIMIZED)
End Function

Function RestoreApp()
   Dim Restoreit%
   Restoreit% = ShowWindow(hWndAccessApp, SW_SHOWNORMAL)
End Function

Public Function UPCGlyph(UPCCode As String) As String

    Select Case Len(UPCCode)
'    Case 6
'        UPCGlyph = UPCCode
'    Case 7
'    Case 8
    Case 10
            UPCGlyph = BuildUPCAGlyph(CheckDigit11("0" & UPCCode))
    Case 11
            UPCGlyph = BuildUPCAGlyph(CheckDigit11(UPCCode))
    Case 12
            UPCGlyph = BuildUPCAGlyph(UPCCode)
    Case Else
            UPCGlyph = ""
    End Select

End Function
