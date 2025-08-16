Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command4_Click()
    Dim rs As Object
    Set rs = Me.RecordsetClone
    Do While Not rs.EOF
        If Not gcfHandleErrors Then Debug.Print rs!Name
        txtFormatter = rs!Name
        Me.Repaint
        If gcfHandleErrors Then On Error GoTo ErrLog
        DoCmd.OpenForm rs!Name
        AddFormats Me.txtFormatter, Application.Forms(rs!Name)
        DoCmd.OpenForm rs!Name, acFormDS
        DoCmd.Save acForm, rs!Name
        DoCmd.Close acForm, rs!Name
Resumen:
        rs.MoveNext
    Loop
Exit Sub
ErrLog:
    Resume Resumen
End Sub
