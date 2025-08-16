Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command2_Click()
Dim strTemp As String
Dim tbdtemp As TableDef
    lstTables = ""
    lstTables1 = ""
    For Each tbdtemp In CurrentDb().TableDefs
        If tbdtemp.Connect = "" Then
            If tbdtemp.RecordCount > 0 Then
'            List0.AddItem tbdtemp.Name
                If Not gcfHandleErrors Then Debug.Print strTemp
                strTemp = Format(CStr(tbdtemp.RecordCount), "@@@@@@@") & " " & tbdtemp.Name
                lstTables.AddItem strTemp
                 lstTables1.AddItem tbdtemp.Name
           End If
        End If
'        Select Case tbdtemp.Connect
 '       Case "ODBC"
 '       Case Else
 '           List0.AddItem tbdtemp.Name & vbtab & tbdtemp.Connect
 '       End Select


    Next tbdtemp
End Sub

Private Sub List7_BeforeUpdate(Cancel As Integer)

End Sub

Private Sub lstTables1_DblClick(Cancel As Integer)
    Dim strTemp As String
    SetOptions False
    strTemp = "Delete *  from " & lstTables1.ItemData(lstTables1.ListIndex)
    DoCmd.RunSQL strTemp, False

    SetOptions True


End Sub
