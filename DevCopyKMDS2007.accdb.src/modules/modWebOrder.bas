Attribute VB_Name = "modWebOrder"
Option Compare Database
Option Explicit
Dim mstrWebDir As String

Public Sub ExportFilesForWeb()
    Dim varReturn As Variant
   SysCmd acSysCmdSetStatus, "Creating Web files..."

    ExportClubFileForWeb
    ExportItemFileForWeb
    ExportItemSalePriceForWeb
    ExportSalesDepartmentFileForWeb
    ExportSizeFileForWeb
    ExportSubcategoryFileForWeb
        varReturn = SysCmd(acSysCmdClearStatus)
        varReturn = SysCmd(acSysCmdSetStatus, " ")
End Sub


Sub ExportItemFileForWeb()

Dim strDesc As String
Dim stDocName As String
Dim stLinkCriteria As String
Dim rstemp As Object
Dim comQDF As Object

Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim strWebTxtFile As String

    If gcfHandleErrors Then On Error GoTo ErrLog

    mstrWebDir = GetKMDSSettings("WebOrdersMailDirectory", "k:\agent" & glnAgentNumber & "\Mail")
    mstrWebDir = Trim(mstrWebDir)
    strWebTxtFile = mstrWebDir & "\Out\Item.txt" ' k:\agent1200\Mail\Out\Item.txt"

    Set rstemp = CreateObject("ADODB.Recordset")
    Set comQDF = CreateObject("ADODB.COMMAND")
    Set txtstrTemp = fsoTemp.OpenTextFile(strWebTxtFile, ForWriting, True, TristateFalse)
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryWebCSVItem"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        Do While Not .EOF
            strDesc = !Description
            strDesc = Replace(strDesc, "'", "")
            strDesc = Replace(strDesc, Chr(34), "")
            txtstrTemp.WriteLine !ID
            txtstrTemp.WriteLine !SizeID
            txtstrTemp.WriteLine !SalesDepartmentID
            txtstrTemp.WriteLine !VendorID
            txtstrTemp.WriteLine !SubCategoryID
            txtstrTemp.WriteLine !CurPrice
            txtstrTemp.WriteLine !Cost
            ' ? make this a system option ?
            'txtstrTemp.WriteLine !Onhand
            txtstrTemp.WriteLine !OnShelf
            txtstrTemp.WriteLine !ChangeDate
            txtstrTemp.WriteLine !OnSale
            txtstrTemp.WriteLine strDesc '!Description
            .MoveNext
        Loop
    End With
    rstemp.Close

ExitSub:
    If gcfHandleErrors Then On Error Resume Next
    txtstrTemp.Close
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ExportItemFileForWeb"
    Resume ExitSub
End Sub

Sub ExportClubFileForWeb()

Dim strName As String
Dim strContact As String
Dim strAddress1 As String
Dim strAddress2 As String
Dim strCity As String
Dim strEmail As String


Dim stDocName As String
Dim stLinkCriteria As String
Dim rstemp As Object
Dim comQDF As Object

Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim strWebTxtFile As String
Dim strTestemail As String

    If gcfHandleErrors Then On Error GoTo ErrLog

    mstrWebDir = GetKMDSSettings("WebOrdersMailDirectory", "k:\agent" & glnAgentNumber & "\Mail")
    mstrWebDir = Trim(mstrWebDir)
    strWebTxtFile = mstrWebDir & "\Out\Club.txt" ' k:\agent1200\Mail\Out\Item.txt"

    Set rstemp = CreateObject("ADODB.Recordset")
    Set comQDF = CreateObject("ADODB.COMMAND")
    Set txtstrTemp = fsoTemp.OpenTextFile(strWebTxtFile, ForWriting, True, TristateFalse)
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryWebCSVClub"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        Do While Not .EOF
            strTestemail = IIf(IsNull(!EMail), "", !EMail)
            'If InStr(strTestemail, "@") > 0 Then

                strName = IIf(IsNull(!Name), "", !Name)
                strName = Replace(strName, "'", "")
                strName = Replace(strName, Chr(34), "")

                strContact = IIf(IsNull(!Contact), "", !Contact)
                strContact = Replace(strContact, "'", "")
                strContact = Replace(strContact, Chr(34), "")

                strAddress1 = IIf(IsNull(!Address1), "", !Address1)
                strAddress1 = Replace(strAddress1, "'", "")
                strAddress1 = Replace(strAddress1, Chr(34), "")

                strAddress2 = IIf(IsNull(!Address2), "", !Address2)
                strAddress2 = Replace(strAddress2, "'", "")
                strAddress2 = Replace(strAddress2, Chr(34), "")

                strCity = IIf(IsNull(!City), "", !City)
                strCity = Replace(strCity, "'", "")
                strCity = Replace(strCity, Chr(34), "")

                strEmail = IIf(IsNull(!EMail), "", !EMail)
                strEmail = Replace(strEmail, "'", "")
                strEmail = Replace(strEmail, Chr(34), "")

                txtstrTemp.WriteLine !ID
                txtstrTemp.WriteLine IIf(IsNull(strName), "", strName)
                txtstrTemp.WriteLine IIf(IsNull(strContact), "", strContact)
                txtstrTemp.WriteLine IIf(IsNull(strAddress1), "", strAddress1)
                txtstrTemp.WriteLine IIf(IsNull(strAddress2), "", strAddress2)
                txtstrTemp.WriteLine IIf(IsNull(strCity), "", strCity)
                txtstrTemp.WriteLine IIf(IsNull(!St), "", !St)
                txtstrTemp.WriteLine IIf(IsNull(!Zip), "", !Zip)
                txtstrTemp.WriteLine IIf(IsNull(!Phone1), "", !Phone1)
                txtstrTemp.WriteLine IIf(IsNull(!Phone2), "", !Phone2)
                txtstrTemp.WriteLine IIf(IsNull(!Phone3), "", !Phone3)
                txtstrTemp.WriteLine IIf(!MaintainParLevels, "-1", "0")
                txtstrTemp.WriteLine IIf(!OrderGuideByCategory, "-1", "0")
                txtstrTemp.WriteLine IIf(IsNull(strEmail), "", strEmail)
                txtstrTemp.WriteLine IIf(IsNull(!Licensee), "", !Licensee)
                txtstrTemp.WriteLine IIf(IsNull(!SetupDate), "", !SetupDate)
                txtstrTemp.WriteLine IIf(IsNull(!ChangeDate), "", !ChangeDate)
                txtstrTemp.WriteLine IIf(IsNull(!OuncesColumn), "", !OuncesColumn)
            'End If
            .MoveNext
        Loop
    End With
    rstemp.Close

ExitSub:
    If gcfHandleErrors Then On Error Resume Next
    txtstrTemp.Close
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ExportClubFileForWeb"
    Resume ExitSub
End Sub

Sub ExportSalesDepartmentFileForWeb()

Dim strDesc As String
Dim stDocName As String
Dim stLinkCriteria As String
Dim rstemp As Object
Dim comQDF As Object

Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim strWebTxtFile As String

    If gcfHandleErrors Then On Error GoTo ErrLog

    mstrWebDir = GetKMDSSettings("WebOrdersMailDirectory", "k:\agent" & glnAgentNumber & "\Mail")
    mstrWebDir = Trim(mstrWebDir)
    strWebTxtFile = mstrWebDir & "\Out\SalesDepartment.txt" ' k:\agent1200\Mail\Out\SalesDepartment.txt"

    Set rstemp = CreateObject("ADODB.Recordset")
    Set comQDF = CreateObject("ADODB.COMMAND")
    Set txtstrTemp = fsoTemp.OpenTextFile(strWebTxtFile, ForWriting, True, TristateFalse)
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryWebCSVSalesDepartment"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        Do While Not .EOF
            strDesc = !Description
            strDesc = Replace(strDesc, "'", "")
            strDesc = Replace(strDesc, Chr(34), "")
            txtstrTemp.WriteLine !ID
            txtstrTemp.WriteLine !SalesGroupID
            txtstrTemp.WriteLine strDesc '!Description
            txtstrTemp.WriteLine !Discounts
            txtstrTemp.WriteLine !AllowClubDiscounts
            txtstrTemp.WriteLine !SetupDate
            txtstrTemp.WriteLine !ChangeDate
            .MoveNext
        Loop
    End With
    rstemp.Close

ExitSub:
    If gcfHandleErrors Then On Error Resume Next
    txtstrTemp.Close
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ExportSalesDepartmentFileForWeb"
    Resume ExitSub

End Sub

Sub ExportSubcategoryFileForWeb()
Dim strDesc As String
Dim stDocName As String
Dim stLinkCriteria As String
Dim rstemp As Object
Dim comQDF As Object

Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim strWebTxtFile As String

    If gcfHandleErrors Then On Error GoTo ErrLog

    mstrWebDir = GetKMDSSettings("WebOrdersMailDirectory", "k:\agent" & glnAgentNumber & "\Mail")
    mstrWebDir = Trim(mstrWebDir)
    strWebTxtFile = mstrWebDir & "\Out\Subcategory.txt" ' k:\agent1200\Mail\Out\SalesDepartment.txt"

    Set rstemp = CreateObject("ADODB.Recordset")
    Set comQDF = CreateObject("ADODB.COMMAND")
    Set txtstrTemp = fsoTemp.OpenTextFile(strWebTxtFile, ForWriting, True, TristateFalse)
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryWebCSVSubcategory"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        Do While Not .EOF
            strDesc = !Description
            strDesc = Replace(strDesc, "'", "")
            strDesc = Replace(strDesc, Chr(34), "")
            txtstrTemp.WriteLine !ID
            txtstrTemp.WriteLine !CategoryID
            txtstrTemp.WriteLine strDesc '!Description
            txtstrTemp.WriteLine !Setup
            txtstrTemp.WriteLine !Change
            .MoveNext
        Loop
    End With
    rstemp.Close

ExitSub:
    If gcfHandleErrors Then On Error Resume Next
    txtstrTemp.Close
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ExportSubcategoryFileForWeb"
    Resume ExitSub


End Sub

Sub ExportItemSalePriceForWeb()
Dim strDesc As String
Dim stDocName As String
Dim stLinkCriteria As String
Dim rstemp As Object
Dim comQDF As Object

Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim strWebTxtFile As String

    If gcfHandleErrors Then On Error GoTo ErrLog

    mstrWebDir = GetKMDSSettings("WebOrdersMailDirectory", "k:\agent" & glnAgentNumber & "\Mail")
    mstrWebDir = Trim(mstrWebDir)
    strWebTxtFile = mstrWebDir & "\Out\ItemSalePrice.txt" ' k:\agent1200\Mail\Out\SalesDepartment.txt"

    Set rstemp = CreateObject("ADODB.Recordset")
    Set comQDF = CreateObject("ADODB.COMMAND")
    Set txtstrTemp = fsoTemp.OpenTextFile(strWebTxtFile, ForWriting, True, TristateFalse)
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryWebCSVSalePrice"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        Do While Not .EOF

            txtstrTemp.WriteLine !ItemID
            txtstrTemp.WriteLine !StartDate
            txtstrTemp.WriteLine !EndDate
            txtstrTemp.WriteLine !OldPrice
            txtstrTemp.WriteLine !SalePrice
            txtstrTemp.WriteLine !SetupDate
            txtstrTemp.WriteLine !ChangeDate
            .MoveNext
        Loop
    End With
    rstemp.Close

ExitSub:
    If gcfHandleErrors Then On Error Resume Next
    txtstrTemp.Close
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ExportSalePriceFileForWeb"
    Resume ExitSub
End Sub

Sub ExportSizeFileForWeb()
Dim strDesc As String
Dim stDocName As String
Dim stLinkCriteria As String
Dim rstemp As Object
Dim comQDF As Object

Dim txtstrTemp As TextStream
Dim fsoTemp As New FileSystemObject
Dim strTemp As String
Dim strWebTxtFile As String

    If gcfHandleErrors Then On Error GoTo ErrLog

    mstrWebDir = GetKMDSSettings("WebOrdersMailDirectory", "k:\agent" & glnAgentNumber & "\Mail")
    mstrWebDir = Trim(mstrWebDir)
    strWebTxtFile = mstrWebDir & "\Out\Size.txt" ' k:\agent1200\Mail\Out\SalesDepartment.txt"

    Set rstemp = CreateObject("ADODB.Recordset")
    Set comQDF = CreateObject("ADODB.COMMAND")
    Set txtstrTemp = fsoTemp.OpenTextFile(strWebTxtFile, ForWriting, True, TristateFalse)
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryWebCSVSize"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        Set rstemp = .Execute
    End With

    With rstemp
        Do While Not .EOF
            strDesc = !Description
            strDesc = Replace(strDesc, "'", "")
            strDesc = Replace(strDesc, Chr(34), "")
            txtstrTemp.WriteLine !ID
            txtstrTemp.WriteLine strDesc '!Description
            txtstrTemp.WriteLine !Oz
            txtstrTemp.WriteLine !CaseQty
            txtstrTemp.WriteLine !SetupDate
            txtstrTemp.WriteLine !ChangeDate
            .MoveNext
        Loop
    End With
    rstemp.Close

ExitSub:
    If gcfHandleErrors Then On Error Resume Next
    txtstrTemp.Close
    Set comQDF = Nothing
    Set rstemp = Nothing
    Set txtstrTemp = Nothing
    Set fsoTemp = Nothing
    Exit Sub

ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, vbInformation, "ExportSizeFileForWeb"
    Resume ExitSub


End Sub
