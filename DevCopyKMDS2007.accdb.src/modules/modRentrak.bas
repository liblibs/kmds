Attribute VB_Name = "modRentrak"
Option Compare Database
Option Explicit

'SELECT Sysfile.AgentNumber, Sysfile.Name, Sysfile.Add1, Sysfile.Add2, Sysfile.Add3, Sysfile.Phone, [CashSysId] & "-" & [ShiftID] & "-" & [InvoHeadID] AS InvoiceNumber, InvoLine5.SetupDate, InvoLine5.Code AS SKU, 0 AS LicenseeID, InvoLine5.InputCode AS UPC, Item.Price AS RetailPrice, InvoLine5.Quantity, [InvoLine5].[Price]/[quantity] AS NetPrice, "n/a" AS Brand, "n/a" AS Distiller
'FROM Sysfile, InvoLine5 INNER JOIN Item ON InvoLine5.Code = Item.ID
'WHERE (((Item.SalesDepartmentID) = 1))
'ORDER BY InvoLine5.CashSysID, InvoLine5.InvoHeadID, InvoLine5.ID;

Public Sub RentrakExport()
Dim fsoTemp As New FileSystemObject
Dim mstrNow As String
Dim strFileNameTemp As String
Dim strFileName As String
Dim strFootName As String
Dim txtstrTemp As TextStream
    If gcfHandleErrors Then On Error GoTo ErrLog
    If gbRentrakInstalled Then
        With fsoTemp
            'this is the way we build a structure
            If Not .FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\Rentrak") Then
                .CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\Rentrak")
            End If
            If Not .FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\Rentrak\Mail") Then
                .CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\Rentrak\Mail")
            End If
            If Not .FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\Rentrak\Mail\Out") Then
                .CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\Rentrak\Mail\Out")
            End If
            If Not .FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\Rentrak\Mail\Out\Archive") Then
                .CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\Rentrak\Mail\Out\Archive")
            End If
            If Not .FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\Rentrak\Mail\In") Then
                .CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\Rentrak\Mail\In")
            End If
            If Not .FolderExists("\\" & gstrCashsysInsLoc & "\KMDS\Rentrak\Mail\In\Archive") Then
                .CreateFolder ("\\" & gstrCashsysInsLoc & "\KMDS\Rentrak\Mail\In\Archive")
            End If
            'this is the way we build a unique file name
            mstrNow = Format(Now, "yyyymmdd_hhnnss")
            strFileName = "\\" & gstrCashsysInsLoc & "\KMDS\Rentrak\Mail\Out\OR_LOC"
            strFileName = strFileName & glnAgentNumber & "_"
            strFileName = strFileName & mstrNow & "_Sales.txt"


            If .FileExists(strFileName) Then
                .DeleteFile strFileName
            End If
            'strFileName = "c:\RE.TXT"

            DoCmd.TransferText acExportDelim, "qryRentrakTransactions Export Specification", "qryRentrakTransactions", strFileName, True

            Set txtstrTemp = .OpenTextFile(strFileName, ForAppending, False)
            '13 9's is the end of file!
            txtstrTemp.WriteLine "9999999999999"
            txtstrTemp.Close
            Set txtstrTemp = Nothing
        End With
    End If
ExitSub:
    Set fsoTemp = Nothing
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "RentrakExport"
    Resume ExitSub
End Sub

Public Sub RentrakRetransmit(BusinessDate As String)

    'Dim comQDF As New Command
    Dim comQDF As Object
    Set comQDF = CreateObject("ADODB.COMMAND")

    If gcfHandleErrors Then On Error GoTo ErrLog
    SetOptions False
    'clear out the work file
    DoCmd.RunSQL "Delete * from Involine5", False
    'append the work file
    With comQDF
        .ActiveConnection = CurrentProject.Connection
        .CommandText = "qryRentrakRetransmitAppend"
        .CommandType = adCmdStoredProc
        'ADO Numeric Datatypes are very particular
        .Parameters.Append .CreateParameter("parTranDate", adDate, adParamInput)
        .Parameters("parTranDate") = CDate(Format(BusinessDate, "mm/dd/yyyy"))
        .Execute
    End With
    RentrakExport
    DoCmd.RunSQL "Delete * from Involine5", False
    SetOptions True
ExitSub:
    'Set fsotemp = Nothing
    Exit Sub
ErrLog:
    MsgBox "Error #" & Err.Number & " " & Err.Description, , "RentrakRetransmit"
    Resume ExitSub

End Sub
