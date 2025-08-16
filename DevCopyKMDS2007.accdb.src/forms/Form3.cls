Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command8_Click()
Dim strInput As String
Dim strOutput As String
Dim nI As Integer
Dim nChr As Integer
Dim strChar As String
Dim bEven As Boolean


    strInput = Trim(txtIn)
    For nI = 1 To Len(strInput)
        If (nI \ 2) * 2 = nI Then
            bEven = True
        Else
            bEven = False
        End If
        strChar = Mid(strInput, nI, 1)
        txtPos = strChar
        nChr = Asc(strChar)
        If bEven Then
            txtEncryptChar = Chr(nChr + nI + 50)
        Else
            txtEncryptChar = Chr(nChr + nI - 30)
        End If
        strOutput = strOutput & txtEncryptChar
        MsgBox (strChar & " " & nChr & " " & bEven)

    Next nI
    txtOut = strOutput
End Sub
