Operation =2
Name ="upcOld"
Option =0
Where ="(((UPCItem.SetupDate)<Now()-30) AND ((qryInvoLineUPC_MRU.MaxOfSetupDate)<Now()-1"
    "80)) OR (((UPCItem.SetupDate)<Now()-30) AND ((qryInvoLineUPC_MRU.MaxOfSetupDate)"
    " Is Null) AND ((Len([id]))>7))"
Begin InputTables
    Name ="UPCItem"
    Name ="qryInvoLineUPC_MRU"
End
Begin OutputColumns
    Expression ="UPCItem.SetupDate"
    Expression ="UPCItem.ID"
    Expression ="UPCItem.ItemID"
    Expression ="qryInvoLineUPC_MRU.MaxOfSetupDate"
    Alias ="ll"
    Expression ="Len([id])"
End
Begin Joins
    LeftTable ="UPCItem"
    RightTable ="qryInvoLineUPC_MRU"
    Expression ="UPCItem.ID = qryInvoLineUPC_MRU.InputCode"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="UPCItem.SetupDate"
        dbInteger "ColumnWidth" ="2310"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ID"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ll"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =207
        Top =-1
        Right =303
        Bottom =98
        Top =0
        Name ="qryInvoLineUPC_MRU"
        Name =""
    End
End
