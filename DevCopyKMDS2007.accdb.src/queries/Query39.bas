Operation =1
Option =0
Begin InputTables
    Name ="Size"
End
Begin OutputColumns
    Expression ="Size.ID"
    Expression ="Size.Description"
    Expression ="Size.Oz"
    Expression ="Size.CaseQty"
    Expression ="Size.SetupDate"
    Expression ="Size.ChangeDate"
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
        dbText "Name" ="Size.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =21
    Top =80
    Right =1003
    Bottom =791
    Left =-1
    Top =-1
    Right =959
    Bottom =441
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Size"
        Name =""
    End
End
