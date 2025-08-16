Attribute VB_Name = "modRedirectHelp"
Option Compare Database
Option Explicit


Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
        (ByVal hwndCaller As Long, ByVal pszFile As String, _
         ByVal uCommand As Long, ByVal dwData As Long) As Long

Const HH_DISPLAY_TOPIC = &H0
Const HH_SET_WIN_TYPE = &H4
Const HH_GET_WIN_TYPE = &H5
Const HH_GET_WIN_HANDLE = &H6
Const HH_DISPLAY_TEXT_POPUP = &HE
Const HH_HELP_CONTEXT = &HF
Const HH_TP_HELP_CONTEXTMENU = &H10
Const HH_TP_HELP_WM_HELP = &H11

Public Sub Show_Help(HelpFileName As String, MycontextID As Long)
    'A specific topic identified by the variable context-ID is started in
    'response to this button click.
    Dim hwndHelp As Long

    'The return value is the window handle of the created Help window.
    Select Case MycontextID
        Case Is = 0
            hwndHelp = HtmlHelp(Application.hWndAccessApp, HelpFileName, _
                       HH_DISPLAY_TOPIC, MycontextID)
        Case Else
            hwndHelp = HtmlHelp(Application.hWndAccessApp, HelpFileName, _
                       HH_HELP_CONTEXT, MycontextID)
    End Select
End Sub

Public Function HelpEntry()
    'Identify the name of the Help file and a possible context-id.
    Dim FormHelpId As Long
    Dim FormHelpFile As String
    Dim curForm As Form

    'Set the curForm variable to the currently active form.
    Set curForm = Screen.ActiveForm

    'As a default, specify a generic Help file and context-id. Note that
    'the location of your file may be different.
    FormHelpId = 10000000

    'Check the Help file property of the form. If a Help file exists,
    'assign the name and context-id to the respective variables.
    If curForm.HelpFile <> "" Then
        'FormHelpFile = Replace(CurrentDb().Name, "KMDS.mdb", "") & "KMDSHelp.chm"
    End If

    'If the Help context-id of the control is not null and greater than
    'zero, assign the value to the variable.
    If Not IsNull(curForm.ActiveControl.Properties("HelpcontextId")) Then
        If curForm.ActiveControl.Properties("HelpcontextId") > 0 Then
            FormHelpId = curForm.ActiveControl.Properties("HelpcontextId")
        Else
            If Not IsNull(curForm.Properties("HelpcontextId")) Then
                FormHelpId = curForm.Properties("HelpcontextId")
            End If
        End If
    End If

    'Call the function to start the Help file, passing it the name of the
    'Help file and context-id.
    Show_Help FormHelpFile, FormHelpId
End Function
