Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmbDescription_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Nz(Me![cmbDescription], 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark


End Sub

Private Sub cmdStep1_SetupCategories_Click()
    SQLPassThru ("SET IDENTITY_INSERT SubCategory ON")
    MsgBox ("Database must be shut down to refresh changes!")
    DoCmd.Quit
End Sub

Private Sub cmdStep2_SetupCategories_Click()
    Dim stDocName As String
If gcfHandleErrors Then On Error GoTo ErrLog
    If MsgBox("This will update the Subcategory records... Continue?", vbYesNo) = vbYes Then
        If MsgBox("Are you sure?", vbYesNo) = vbYes Then


            SetOptions False
            stDocName = "qrySetupSubCategory_Step_1_AppendCategory"
            DoCmd.OpenQuery stDocName, acNormal, acEdit
            stDocName = "qrySetupSubCategory_Step_2_AppendSubCategory_0"
            DoCmd.OpenQuery stDocName, acNormal, acEdit

            stDocName = "qrySetupSubCategory_Step_3_UpdateSubCategory_0"
            DoCmd.OpenQuery stDocName, acNormal, acEdit
            SQLPassThru ("SET IDENTITY_INSERT SubCategory ON")
            'SQLPassThru ("sp_SubCategoryIDENTITY_INSERT_OFF")
            SetOptions True

            MsgBox ("Database must be shut down to refresh changes!")
            DoCmd.Quit

        End If
    End If

ExitSub:
    Exit Sub

ErrLog:
    MsgBox Err.Description
    Resume ExitSub


End Sub

Private Sub Command17_Click()

End Sub

Private Sub cmdStep3_SetupOLCCCategories_Click()
   Dim stDocName As String
If gcfHandleErrors Then On Error GoTo Err_cmdSetupCategories_Click
    If MsgBox("This will update the Subcategory records... Continue?", vbYesNo) = vbYes Then
        If MsgBox("Are you sure?", vbYesNo) = vbYes Then



            SetOptions False
            stDocName = "qrySetupSubCategory_Step_4_AppendSubCategory"
            DoCmd.OpenQuery stDocName, acNormal, acEdit
            stDocName = "qrySetupSubCategory_Step_5_UpdateItemSubCategory"
            DoCmd.OpenQuery stDocName, acNormal, acEdit
            SetOptions True
        End If
    End If

Exit_cmdSetupCategories_Click:
    Exit Sub

Err_cmdSetupCategories_Click:
    MsgBox Err.Description
    Resume Exit_cmdSetupCategories_Click

End Sub

Private Sub Form_Current()
    cmbDescription = ID
End Sub

Private Sub Form_Load()
    Me.InsideHeight = Box11.Height + FormFooter.Height + 100
    Me.InsideWidth = Box11.Width + 500
    cmbDescription.ColumnHidden = True
    Setup.ColumnHidden = True
    Change.ColumnHidden = True

End Sub

Private Sub cmdSetupCategories_Click()

End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    Call DoMouseWheel(Me, Count)

End Sub

Private Sub Form_Open(Cancel As Integer)
    Cancel = FormLogin(glnCashierID, Me)
End Sub

Private Sub Form_Close()
    FormLogout glnCashierID, Me
End Sub
