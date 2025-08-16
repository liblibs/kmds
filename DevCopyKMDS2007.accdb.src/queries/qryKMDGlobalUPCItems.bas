Operation =1
Option =0
Where ="(((Len(UPCItem.id))>5) And ((Item.SalesDepartmentID)=1))"
Begin InputTables
    Name ="Sysfile"
    Name ="UPCItem"
    Name ="Item"
End
Begin OutputColumns
    Expression ="UPCItem.ID"
    Expression ="UPCItem.ItemID"
    Expression ="Sysfile.AgentNumber"
    Alias ="lid"
    Expression ="Len([UPCItem].[id])"
    Expression ="UPCItem.ChangeDate"
End
Begin Joins
    LeftTable ="UPCItem"
    RightTable ="Item"
    Expression ="UPCItem.ItemID = Item.ID"
    Flag =1
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
        dbText "Name" ="UPCItem.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.AgentNumber"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ChangeDate"
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
    Bottom =315
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Item"
        Name =""
    End
End
