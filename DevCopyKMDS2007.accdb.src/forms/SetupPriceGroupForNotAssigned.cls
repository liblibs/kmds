Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdStep1_SetupCategories_Click()
    SQLPassThru ("SET IDENTITY_INSERT PriceGroup ON")
    MsgBox ("Database must be shut down to refresh changes!")
    DoCmd.Quit
End Sub

Private Sub cmdStep2_SetupCategories_Click()
    Dim stDocName As String
If gcfHandleErrors Then On Error GoTo ErrLog
    'If MsgBox("This will update the Subcategory records... Continue?", vbYesNo) = vbYes Then
     '   If MsgBox("Are you sure?", vbYesNo) = vbYes Then


            SetOptions False
            stDocName = "qrySetupPriceGroup_Step_1_PriceGroup0"
            DoCmd.OpenQuery stDocName, acNormal, acEdit
            stDocName = "qrySetupPriceGroup_Step_2_UpdatePriceGroup_0"
            DoCmd.OpenQuery stDocName, acNormal, acEdit
            SQLPassThru ("SET IDENTITY_INSERT SubCategory ON")
            'SQLPassThru ("sp_SubCategoryIDENTITY_INSERT_OFF")
            SetOptions True

            MsgBox ("Database must be shut down to refresh changes!")
            DoCmd.Quit



ExitSub:
    Exit Sub

ErrLog:
    MsgBox Err.Description
    Resume ExitSub


End Sub
