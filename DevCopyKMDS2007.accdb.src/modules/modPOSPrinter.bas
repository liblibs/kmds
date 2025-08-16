Attribute VB_Name = "modPOSPrinter"
Option Compare Database




Option Explicit

Private Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
   "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
    ByVal pDefault As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias _
   "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
   pDocInfo As DOCINFO) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
   pcWritten As Long) As Long

Public Sub HardCopy(strOut As String, Optional DocDescription As String)

Dim lhPrinter As Long
Dim lpcWritten As Long
Dim sWrittenData As String

Dim lReturn As Long
Dim lDoc As Long
Dim MyDocInfo As DOCINFO

    If gcfHandleErrors Then On Error GoTo ErrLog
    If IsNull(gstrPOSPrinterPort) Then Exit Sub

    lReturn = OpenPrinter(gstrPOSPrinterPort, lhPrinter, 0)
    If lReturn = 0 Then
        MsgBox "The Printer named '" & gstrPOSPrinterPort & "' is not found on this machine"

        Exit Sub
    End If
    MyDocInfo.pDocName = "Invoice " & DocDescription
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)

    Call StartPagePrinter(lhPrinter)
    lReturn = WritePrinter(lhPrinter, ByVal strOut, Len(strOut), lpcWritten)
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)


ExitSub:
    Exit Sub
ErrLog:
    fMsgBox "HardCopy"
    Resume ExitSub
End Sub

Public Sub FaxCopy(ReportName As String, FaxNumber As String, Optional Displayname As String, Optional RecipientName As String, Optional CoverpageSubject As String)

    Dim faxserver As Object
    Dim faxdoc As Object

    Set faxserver = CreateObject("FaxServer.FaxServer")
    faxserver.Connect ("")
    Set faxdoc = faxserver.CreateDocument(ReportName)
    With faxdoc
        .Displayname = Displayname
        .FaxNumber = FaxNumber
        .RecipientName = RecipientName
        .CoverpageSubject = CoverpageSubject
        .Send
    End With


ExitSub:
    Set faxdoc = Nothing
    Set faxserver = Nothing
    Exit Sub
ErrLog:
    fMsgBox "FaxCopy"
    Resume ExitSub
End Sub
