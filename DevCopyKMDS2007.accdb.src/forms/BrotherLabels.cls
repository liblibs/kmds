Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command2_Click()

'Brother developer center... Labelling SDK (b-PAC)
'http://www.brother.com/product/dev/label/bpac/features/index.htm

'redist files com ponents are located at
'ftp.kmds.us/Agents/BrotherLabel/Redist

    If gcfHandleErrors Then On Error GoTo ErrLog
    On Error GoTo ErrLog

    Dim strLBXpath As String
    Dim strqryName As String

    Dim strID As String
    Dim strDescr As String
    Dim strCurPrice As String
    Dim strSalesDpt As String
    Dim strPOZ As String
    Dim strBarCode As String
    Dim strSizeD As String

    Dim doc As bpac.Document
    Dim rstemp As Recordset
    Dim strText As String


    strqryName = "qryItemLabelsForBrother"
    strLBXpath = GetKMDSSettings("BrotherLabel", "K:\Brother\DK205 KMDS Shelf Tag.lbx")


    Set doc = CreateObject("bpac.Document")
    ' Check that the printer is online
    If doc.Printer.IsPrinterOnline(doc.GetPrinterName) Then
        Debug.Print "Online"
'        MsgBox ("Printer is online")
    Else
        Debug.Print "Offline"
        MsgBox (doc.Printer.Name & " on port " & doc.Printer.PortName & "  is off line... try again!")
        GoTo ExitSub
    End If


    '1. open the grid as a recordset
    Set rstemp = CurrentDb().OpenRecordset(strqryName)

    With rstemp
        If .RecordCount > 0 Then
             ' Open template
            If Not doc.Open(strLBXpath) Then
                MsgBox "Can not open Label template file named: " & vbCrLf & strLBXpath
                GoTo ExitSub
            End If
            ' Print setting start
            If Not doc.StartPrint("", bpac.PrintOptionConstants.bpoAutoCut) Then
                MsgBox "Can not StartPrint for Printername named : " & vbCrLf & doc.GetPrinterName
                GoTo ExitSub
            End If
            .MoveLast
            If MsgBox(.RecordCount & " labels to be printed", vbOKCancel) = vbCancel Then
                GoTo ExitSub
            End If
            .MoveFirst

            .MoveFirst
            Do While Not .EOF
                doc.GetObject("objID").Text = !ID
                doc.GetObject("objDescr").Text = !DESCR
                doc.GetObject("objCurPrice").Text = !CurPrice
                doc.GetObject("objSalesDpt").Text = !SalesDpt
                doc.GetObject("objPOZ").Text = !POZ
                doc.GetObject("objBarCode").Text = !BarCode
                doc.GetObject("objSizeD").Text = !SizeD
                ' Adds a print job
                doc.PrintOut 1, bpac.PrintOptionConstants.bpoDefault
                .MoveNext
            Loop
            ' Print setting end (=Print start)
            doc.EndPrint
            ' Close template
            doc.Close
        Else
            MsgBox "No Labels selected to print... try again"
        End If
    End With
ExitSub:
    Exit Sub

ErrLog:
    strText = strLBXpath & vbCrLf & vbCrLf & "Verify the following Object names in the above Label template!" & vbCrLf
    strText = strText & "objID" & vbCrLf
    strText = strText & "objDescr" & vbCrLf
    strText = strText & "objCurPrice" & vbCrLf
    strText = strText & "objSalesDpt" & vbCrLf
    strText = strText & "objPOZ" & vbCrLf
    strText = strText & "objBarCode" & vbCrLf
    strText = strText & "objSizeD" & vbCrLf
    MsgBox "Error: " & Err.Number & " " & Err.Description & " Brother Label Print", vbRed
    MsgBox strText
    Resume ExitSub
 End Sub
