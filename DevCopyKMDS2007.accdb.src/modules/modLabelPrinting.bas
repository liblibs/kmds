Attribute VB_Name = "modLabelPrinting"



'*********************************************************
    'Declarations section of the module.
 '**********************************************************



 Option Compare Database
 Option Explicit

 Dim LabelBlanks&
 Dim LabelCopies&
 Dim BlankCount&
 Dim CopyCount&

 '==========================================================
    ' The following function will cause an input box to
    ' display when the report is run that prompts the user
    ' for the number of used labels to skip and how many
    ' copies of each label should be printed.
 '===========================================================

 Function LabelSetup()
 LabelBlanks& = Val(InputBox$("Enter Number of blank labels to skip"))
 LabelCopies& = Val(InputBox$("Enter Number of Copies to Print"))
 If LabelBlanks& < 0 Then LabelBlanks& = 0
 If LabelCopies& < 1 Then LabelCopies& = 1
 End Function

 '===========================================================
    ' The following function sets the variables to a zero
 '===========================================================

 Function LabelInitialize()
    BlankCount& = 0
    CopyCount& = 0
 End Function

 '===========================================================
    ' The following function is the main part of this code
    ' that allows the labels to print as the user desires.
 '===========================================================

 Function LabelLayout(r As Report)
    If BlankCount& < LabelBlanks& Then
       r.NextRecord = False
       r.PrintSection = False
       BlankCount& = BlankCount& + 1
    Else
       If CopyCount& < (LabelCopies& - 1) Then
          r.NextRecord = False
          CopyCount& = CopyCount& + 1
       Else
          CopyCount& = 0
       End If
    End If
 End Function
