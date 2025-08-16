Attribute VB_Name = "modConditionalFormats"
Option Compare Database
Option Explicit


Function AddFormats(ctlSource As Control, frm As Form) As Integer
    Dim frmTemp As Form
    Dim ctl As Control
    Dim fcdSource As FormatCondition
    Dim fcdDestination As FormatCondition
    Dim varOperator As Variant
    Dim varType As Variant
    Dim varExpression1 As Variant
    Dim varExpression2 As Variant
    Dim intConditionCount As Integer
    Dim intCount As Integer

    If gcfHandleErrors Then On Error GoTo ErrLog

    intConditionCount = ctlSource.FormatConditions.Count
    For Each frmTemp In Application.Forms
        If frmTemp.Name <> frm.Name Then
            For Each ctl In frmTemp.Controls
                If ctl.Name = ctlSource.Name Then
                    ' This is the source.  Don't apply formatting.
                ElseIf ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then
                    intCount = 0

                    ' Bulk remove all current FormatConditions
                    ctl.FormatConditions.Delete

                    Do Until intCount = intConditionCount
                        Set fcdSource = ctlSource.FormatConditions.Item(intCount)

                        varOperator = fcdSource.Operator
                        varType = fcdSource.Type
                        varExpression1 = fcdSource.Expression1
                        varExpression2 = fcdSource.Expression2

                        ' Add the FormatCondition
                        ctl.FormatConditions.Add varType, varOperator, varExpression1, varExpression2

                        ' Reference the FormatCondition to apply formatting.
                        ' Note: The FormatCondition cannot be referenced
                        ' in this manner until it exists.
                        Set fcdDestination = ctl.FormatConditions.Item(intCount)

                        With fcdDestination
                            .BackColor = fcdSource.BackColor
                            .FontBold = fcdSource.FontBold
                            .FontItalic = fcdSource.FontItalic
                            .FontUnderline = fcdSource.FontUnderline
                            .ForeColor = fcdSource.ForeColor
                        End With

                        ' Move to the next FormatCondition
                        intCount = intCount + 1

                    Loop
                End If
            Next ctl
        End If
    Next frmTemp
    ' Cleanup
    AddFormats = intConditionCount
    If Not gcfHandleErrors Then Debug.Print "There were " & AddFormats & " Conditional Format(s) applied to all text and combo boxes except the source."
    Set ctl = Nothing
    Set fcdSource = Nothing
    Set fcdDestination = Nothing
    Set varOperator = Nothing
    Set varType = Nothing
    Set varExpression1 = Nothing
    Set varExpression2 = Nothing
    intConditionCount = 0
    intCount = 0
ExitSub:
    Exit Function
ErrLog:
    fMsgBox "AddFormats"
    Resume ExitSub
End Function
