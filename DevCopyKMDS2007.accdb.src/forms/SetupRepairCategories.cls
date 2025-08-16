Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdStep1_SetupCategories_Click()
    SQLPassThru ("sp_Setup_ID_0_SubCategory_Vendor_PriceGroup")
    MsgBox ("Database must be shut down to reflect changes!")
    DoCmd.Quit
End Sub

Private Sub cmdStep2_SetupCategories_Click()
    Dim stDocName As String
If gcfHandleErrors Then On Error GoTo ErrLog
    If MsgBox("This will update the Subcategory records... Continue?", vbYesNo) = vbYes Then
        If MsgBox("Are you sure?", vbYesNo) = vbYes Then


            SetOptions False

            stDocName = "qrySetupSubCategory_Step_3_UpdateSubCategory_0"
            DoCmd.OpenQuery stDocName, acNormal, acEdit
            stDocName = "qrySetupVendor_Step_2_UpdateVendor_0"
            DoCmd.OpenQuery stDocName, acNormal, acEdit
            stDocName = "qrySetupPriceGroup_Step_2_UpdatePriceGroup_0"
            DoCmd.OpenQuery stDocName, acNormal, acEdit

            SetOptions True

            MsgBox ("finished")

        End If
    End If

ExitSub:
    Exit Sub

ErrLog:
    MsgBox Err.Description
    Resume ExitSub


End Sub


Private Sub cmdStep3_SetupOLCCCategories_Click()
   Dim stDocName As String
If gcfHandleErrors Then On Error GoTo Err_cmdSetupCategories_Click
    If MsgBox("This will update the Subcategory records... Continue?", vbYesNo) = vbYes Then
        If MsgBox("Are you sure?", vbYesNo) = vbYes Then



            SetOptions False
'            stDocName = "qrySetupSubCategory_Step_4_AppendSubCategory"
'            DoCmd.OpenQuery stDocName, acNormal, acEdit
'            stDocName = "qrySetupSubCategory_Step_5_UpdateItemSubCategory"
'            DoCmd.OpenQuery stDocName, acNormal, acEdit

            stDocName = "qryItemsSetCategoryToOLCC_Step_1_Reset_To_0"
            DoCmd.OpenQuery stDocName, acNormal, acEdit
            stDocName = "qryItemsSetCategoryToOLCC_Step_2_Delete_Liquor_SubCategory"
            DoCmd.OpenQuery stDocName, acNormal, acEdit
            stDocName = "qryItemsSetCategoryToOLCC_Step_3_Append_SubCategories"
            DoCmd.OpenQuery stDocName, acNormal, acEdit
            stDocName = "qryItemsSetCategoryToOLCC_Step_4_Update_SubCategoryID"
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
