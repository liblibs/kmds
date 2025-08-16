Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




'*********************************************************
    'Declarations section of the module.
 '**********************************************************



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
 End Function

 '===========================================================
    ' The following function is the main part of this code
    ' that allows the labels to print as the user desires.
 '===========================================================

 Function LabelLayout(r As Report)
 End Function



Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
'    MsgBox Me.Labels
    If BlankCount& < LabelBlanks& Then
       Me.NextRecord = False
       Me.PrintSection = False
       BlankCount& = BlankCount& + 1
    Else
       If CopyCount& < (Me.Labels - 1) Then
          Me.NextRecord = False
          CopyCount& = CopyCount& + 1
       Else
          CopyCount& = 0
       End If
    End If

End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
    MsgBox "Dataerr " & DataErr
End Sub

Private Sub Report_Open(Cancel As Integer)
 LabelBlanks& = Val(InputBox$("Enter Number of blank labels to skip"))
 'LabelCopies& = Val(InputBox$("Enter Number of Copies to Print"))
 If LabelBlanks& < 0 Then LabelBlanks& = 0
 If LabelCopies& < 1 Then LabelCopies& = 1

End Sub


Private Sub ReportFooter_Format(Cancel As Integer, FormatCount As Integer)
    BlankCount& = 0

End Sub
