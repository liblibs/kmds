Attribute VB_Name = "zTest"
Option Compare Database
Option Explicit

Function TestDAO()
    Dim rsDAOLocal As Recordset
    Dim rsDAOSQL As Recordset
    Dim db As Database
    Dim strSQL As String

    Set db = CurrentDb

    'Access query
    strSQL = "SELECT * FROM qryCreatePOBasedOnSalesHistory"
    Set rsDAOLocal = db.OpenRecordset(strSQL, dbOpenDynaset)
    If rsDAOLocal.RecordCount > 0 Then
        MsgBox "Local Tables - DAO connection successful"
    End If

    'Access table
    strSQL = "SELECT * FROM AdjustReasons"
    Set rsDAOLocal = db.OpenRecordset(strSQL, dbOpenDynaset)
    If rsDAOLocal.RecordCount > 0 Then
        MsgBox "Local Tables - DAO connection successful"
    End If

    'SQL table
    strSQL = "SELECT * FROM Agents"
    Set rsDAOLocal = db.OpenRecordset(strSQL, dbOpenDynaset)
    If rsDAOLocal.RecordCount > 0 Then
        MsgBox "SQL Tables - DAO connection successful"
    End If

    'Note to self: No need to use DAO, since we already have ACE
End Function

Function TestCheckParam()
    Dim rsDAOLocal As DAO.Recordset
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim prm As DAO.Parameter
    Dim strSQL As String

    Set db = CurrentDb

    ' Access query
    strSQL = "SELECT * FROM qryCreatePOBasedOnSalesHistory"

    ' Create a QueryDef object
    Set qdf = db.CreateQueryDef("", strSQL)

    ' Loop through each parameter in the query and set its value
    For Each prm In qdf.Parameters
        prm.Value = Eval(prm.Name)
    Next prm

    ' Open the recordset
    Set rsDAOLocal = qdf.OpenRecordset(dbOpenDynaset)

    If rsDAOLocal.RecordCount > 0 Then
        MsgBox "Local Tables - DAO connection successful"
    End If

    ' Clean up
    rsDAOLocal.Close
    Set rsDAOLocal = Nothing
    Set qdf = Nothing
    Set db = Nothing
End Function
