Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"X\"))"
Begin InputTables
    Name ="InvoLine"
    Name ="Agents"
    Name ="Item"
End
Begin OutputColumns
    Alias ="Store Number"
    Expression ="InvoLine.SalesDept"
    Alias ="PostDate"
    Expression ="InvoLine.SetupDate"
    Alias ="StoreName"
    Expression ="Agents.LastName"
    Alias ="TransferNumber"
    Expression ="InvoLine.SalesGroup"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.ID"
    Expression ="InvoLine.Code"
    Expression ="Item.Description"
    Expression ="InvoLine.Quantity"
    Expression ="InvoLine.Cost"
    Expression ="InvoLine.Price"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Agents"
    Expression ="InvoLine.SalesDept = Agents.AgentNumber"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine.SalesDept"
    Flag =0
    Expression ="InvoLine.SetupDate"
    Flag =1
    Expression ="InvoLine.ID"
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
        dbText "Name" ="PostDate"
        dbInteger "ColumnWidth" ="2310"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StoreName"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransferNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Store Number"
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
    Bottom =233
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Agents"
        Name =""
    End
    Begin
        Left =407
        Top =-5
        Right =503
        Bottom =109
        Top =0
        Name ="Item"
        Name =""
    End
End
