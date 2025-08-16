Operation =1
Option =0
Where ="(((MSysObjects.Name)<>\"qryForms\") AND ((MSysObjects.Type)=-32768))"
Begin InputTables
    Name ="MSysObjects"
End
Begin OutputColumns
    Expression ="MSysObjects.Name"
    Expression ="MSysObjects.Type"
End
Begin OrderBy
    Expression ="MSysObjects.Name"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="MSysObjects.Name"
        dbInteger "ColumnWidth" ="4395"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1225
    Bottom =677
    Left =-1
    Top =-1
    Right =1209
    Bottom =270
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =183
        Bottom =255
        Top =0
        Name ="MSysObjects"
        Name =""
    End
End
