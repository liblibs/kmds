Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

Private Sub Command7_Click()
'Dim tbdTemp As TableDef
'
'   Dim dbsNorthwind As Database
'   Dim tdfNew As TableDef
'   Dim tdfLoop As TableDef
'   Dim prpLoop As Property

'For Each tbdTemp In CurrentDb().TableDefs
'    If InStr(tbdTemp.Connect, "ODBC") > 0 Then
'        If not gcfHandleErrors then Debug.Print tbdTemp.Name & vbtab & vbtab & tbdTemp.Connect
'        tbdTemp.Connect = txtNewConnectString
'        tbdTemp.RefreshLink
'
'
'    End If
'Next

    Dim strDatabaseName As String

   Dim wrkMain As Workspace
   Dim dbsPubs As Database
   Dim rsInvoLine As DAO.Recordset
   Dim rsInvoline5 As DAO.Recordset
    Dim nI As Integer
   ' Create Microsoft Jet Workspace object.
   Dim varDir As Variant
   Set wrkMain = CreateWorkspace("", "admin", "", dbUseJet)

   ' Open a Database object based on information in
   ' the connect string.

   'Note: The DSN referenced below must be configured to
   '       use Microsoft Windows NT Authentication Mode to
   '       authorize user access to the Microsoft SQL Server.
'    , dbDriverNoPrompt, False, _
'        "ODBC;DATABASE=pubs;DSN=Publishers")'

   ' Enumerate the Properties collection of the Database
   ' object.
If gcfHandleErrors Then On Error GoTo ErrLog

varDir = Dir(txtDir & "CR*.mdb")
Do While varDir <> ""
    txtCurFile = txtDir & varDir
    Me.Repaint
    Set dbsPubs = wrkMain.OpenDatabase(txtCurFile)
    With dbsPubs
        Set rsInvoline5 = CurrentDb().OpenRecordset("InvoiceLineLD")
        Set rsInvoLine = dbsPubs.OpenRecordset("InvoLine")
        With rsInvoLine
            Do While Not .EOF
                rsInvoline5.AddNew
                For nI = 0 To rsInvoline5.Fields.Count - 1
                    rsInvoline5.Fields(nI) = .Fields(nI)
                Next nI
                rsInvoline5.Update

                If Not gcfHandleErrors Then Debug.Print rsInvoLine!InvoHeadID

                .MoveNext
            Loop
            .Close
            rsInvoline5.Close
        End With
        Set rsInvoLine = Nothing
        Set rsInvoline5 = Nothing

    End With
    dbsPubs.Close
    Set dbsPubs = Nothing
    If Not gcfHandleErrors Then Debug.Print txtDir & varDir
    varDir = Dir
Loop

   wrkMain.Close

ExitSub:
    Exit Sub
ErrLog:
    Select Case Err.Number
    Case 3022

        Resume Next
    Case Else
        MsgBox ("error #" & Err.Number)
        Resume ExitSub
    End Select

End Sub
