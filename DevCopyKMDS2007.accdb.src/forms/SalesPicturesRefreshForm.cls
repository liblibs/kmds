Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdChangeName_Click()
Dim strNew As String
Dim fso As FileSystemObject
Dim f As Folder

Dim stDocName As String
Dim strSQL  As String

'    'get the Original SalesPictures.ID
'    'to update the Items with this ID with the new ID
'    Set RS = Me.Recordset.Clone
'    RS.FindFirst "[FileName] = '" & lstPictures & "'"
'    If Not RS.EOF Then Me.Bookmark = RS.Bookmark
'    lngSalesPictureID_Original = ID


    If gcfHandleErrors Then On Error GoTo ErrLog
    Set fso = New FileSystemObject

    strNew = Replace(txtFileName, txtOldName, txtNewName)

    If fso.FileExists(strNew) Then
        MsgBox strNew & " already exists"

        Exit Sub
        fso.DeleteFile strNew
    End If

    fso.CopyFile txtFileName, strNew, True
    fso.DeleteFile txtFileName, True

    'Change the file name in the data file..
    strSQL = "Update SalesPictures SET SalesPictures.FileName ='" & strNew & "'"
    strSQL = strSQL & " WHERE SalesPictures.FileName = '" & txtFileName & "'"
    SetOptions False
    DoCmd.RunSQL strSQL, True
    SetOptions True

    RefreshListbox

    'get the Replace  SalesPictures.ID
    'to update the Items with the lngSalesPictureID_Original with the lngSalesPictureID_Replace
    Me.Requery

    lstPictures.SetFocus
    cmdChangeName.Visible = False

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "cmdChangeName_Click"
    Resume ExitSub
End Sub

Private Sub Form_Load()
    If gcfHandleErrors Then On Error GoTo ErrLog
    Me.InsideWidth = boxScreen.Width + 500
    Me.InsideHeight = Detail.Height + Me.FormFooter.Height '+ 500
    RefreshListbox
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Form_Load"
    Resume ExitSub

End Sub

Private Sub lstPictures_AfterUpdate()
    Dim rs As Object
    If gcfHandleErrors Then On Error GoTo ErrLog
    'display the picture and store the current file name
    ' in case user wants to change picture name
    txtFileName = lstPictures
    'remove the directory and extension file
    txtOldName = Replace(txtFileName, txtPictureDirectory & "\", "")
    txtOldName = Replace(txtOldName, ".jpg", "")
    txtNewName = txtOldName
    imgPicture.Picture = txtFileName
    imgPicture.Visible = True
    cmdChangeName.Visible = False
    ' Find the record that matches the control.

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[FileName] = '" & txtFileName & "'"
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "Form_Load"
    Resume ExitSub
End Sub


Private Sub RefreshListbox()
Dim strDestError As String
Dim txtFileName As String
Dim strSQL As String
Dim strTemp As String
If gcfHandleErrors Then On Error GoTo ErrLog
    txtServer = gstrCashsysInsLoc
    txtPictureDirectory = "\\" & gstrCashsysInsLoc & "\KMDS\Pictures"
    lstPictures.RowSource = ""
    SetOptions False
    '1. delete tempory sales pictures rows
    strSQL = "Delete * from SalesPicturesTemp"
    DoCmd.RunSQL strSQL

    '2. retrieve listing of *.jpg files in directory and insert as rows in temp table
    strTemp = Dir(txtPictureDirectory & "\*.JPG")
    Do While strTemp <> vbNullString
        txtFileName = (txtPictureDirectory & "\" & strTemp)
        lstPictures.AddItem txtFileName, lstPictures.ListCount
        strSQL = "INSERT INTO SalesPicturesTemp ( FileName ) SELECT '" & txtFileName & "'"
        On Error Resume Next
        DoCmd.RunSQL strSQL
        strTemp = Dir
    Loop
    '3. Append New Pictures into SalesPictures Table
    DoCmd.OpenQuery "qrySalesPictures_1_Append"
    '4  Delete any rows not found in Temp table i.e., deleted pictures
    DoCmd.OpenQuery "qrySalesPictures_2_Delete"
    '5  reset any broken Item.PicturesID links to 0
    DoCmd.OpenQuery "qrySalesPictures_3_ResetItemSalespictureID"

    SetOptions True

ExitSub:
    Exit Sub

ErrLog:
    EMsgBox Me, "RefreshListbox"
    Resume ExitSub
End Sub



Private Sub txtNewName_Change()
    cmdChangeName.Visible = True
End Sub
'------------------------------------------------------------
' cmdMailPictures_Click
'
'------------------------------------------------------------
Private Sub cmdMailPictures_Click()
On Error GoTo cmdMailPictures_Click_Err

    On Error Resume Next
    MailPictures False


cmdMailPictures_Click_Exit:
    Exit Sub

cmdMailPictures_Click_Err:
    MsgBox Error$
    Resume cmdMailPictures_Click_Exit

End Sub
