Attribute VB_Name = "modBrotherLabelPrinter"
Option Compare Database
Option Explicit

 Public Sub BrotherLabels(ShelfOrItem As String)

'Brother developer center... Labelling SDK (b-PAC)
'http://www.brother.com/product/dev/label/bpac/features/index.htm

'redist files com ponents are located at
'ftp.kmds.us/Agents/BrotherLabel/Redist

    If gcfHandleErrors Then On Error GoTo ErrLog
    On Error GoTo ErrLog

    Dim strLBXpath As String
    Dim strLBXpathLiquor As String
    Dim strLBXpathLiquorNot As String

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
    Dim nNumberOfCopies As Integer

    Dim bLiquor As Boolean

    strqryName = "qryItemLabelsForBrother"
    Select Case UCase(ShelfOrItem)
    Case "SHELF"
        strLBXpathLiquorNot = GetKMDSSettings("BrotherLabelShelf", "K:\Brother\DK205 KMDS Shelf Tag.lbx")
        strLBXpathLiquor = GetKMDSSettings("BrotherLabelShelfOLCC", "K:\Brother\DK205 KMDS OLCC Clone Shelf Tag.lbx")
    Case "ITEM"
        strLBXpath = GetKMDSSettings("BrotherLabelItem", "K:\Brother\DK205 KMDS Item Tag.lbx")
    Case Else
        MsgBox UCase(ShelfOrItem) & " Not DEFINED", "BrotherLabels"
        Exit Sub
    End Select


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
            .MoveLast
            If .RecordCount > 1 Then
                If MsgBox(.RecordCount & " labels to be printed", vbOKCancel) = vbCancel Then
                    GoTo ExitSub
                End If
            Else
                If !Labels > 1 Then
                    If MsgBox(!Labels & " labels to be printed", vbOKCancel) = vbCancel Then
                        GoTo ExitSub
                    End If
                End If
            End If

            .MoveFirst
            Do While Not .EOF
                Select Case UCase(ShelfOrItem)
                Case "SHELF"
                    bLiquor = False
                    If !ID > 999 And !ID < 100000 Then bLiquor = True

                    If bLiquor Then
                        strLBXpath = strLBXpathLiquor
                    Else
                        strLBXpath = strLBXpathLiquorNot
                    End If

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

                    doc.GetObject("objID").Text = !ID
                    doc.GetObject("objDescr").Text = !DESCR
                    doc.GetObject("objCurPrice").Text = Format(!CurPrice, "Currency")
                    doc.GetObject("objBarCode").Text = !BarCode
                    doc.GetObject("objPOZ").Text = !POZ
                    If Not bLiquor Then
                        doc.GetObject("objSalesDpt").Text = !SalesDpt
                        doc.GetObject("objSizeD").Text = !SizeD
                    End If
                    nNumberOfCopies = 1
                    ' Adds a print job
                Case "ITEM"
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
                    'doc.GetObject("objID").Text = !ID
                    doc.GetObject("objDescr").Text = !DESCR
                    doc.GetObject("objCurPrice").Text = Format(!CurPrice, "Currency")
                    'doc.GetObject("objSalesDpt").Text = !SalesDpt
                    'doc.GetObject("objPOZ").Text = !POZ
                    doc.GetObject("objBarCode").Text = !BarCode
                    'doc.GetObject("objSizeD").Text = !SizeD
                    nNumberOfCopies = !Labels
                End Select
                doc.PrintOut nNumberOfCopies, bpac.PrintOptionConstants.bpoDefault
                .MoveNext
                ' Print setting end (=Print start)
                doc.EndPrint
                ' Close template
                doc.Close
            Loop
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
