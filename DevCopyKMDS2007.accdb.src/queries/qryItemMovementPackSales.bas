Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"S\"))"
Begin InputTables
    Name ="Item"
    Name ="UPCItem"
    Name ="InvoLine"
End
Begin OutputColumns
    Alias ="CartonID"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Alias ="PackID"
    Expression ="UPCItem.ItemID"
    Alias ="PacksPerCarton"
    Expression ="IIf(IsNull([UnitsPerPack]) Or [UnitsPerPack]=0,1,[UnitsPerPack])"
    Alias ="PacksSold"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="PackDollars"
    Expression ="Sum(InvoLine.Price)"
    Alias ="MMYY"
    Expression ="CDate(Format(InvoLine.SetupDate,\"mm/dd/yyyy\"))"
    Alias ="MM"
    Expression ="Format(InvoLine.SetupDate,\"mm\")"
    Alias ="YY"
    Expression ="Format(InvoLine.SetupDate,\"yyyy\")"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="UPCItem"
    Expression ="Item.SubPack = UPCItem.ID"
    Flag =1
    LeftTable ="UPCItem"
    RightTable ="InvoLine"
    Expression ="UPCItem.ItemID = InvoLine.Code"
    Flag =1
End
Begin Groups
    Expression ="Item.ID"
    GroupLevel =0
    Expression ="Item.Description"
    GroupLevel =0
    Expression ="UPCItem.ItemID"
    GroupLevel =0
    Expression ="IIf(IsNull([UnitsPerPack]) Or [UnitsPerPack]=0,1,[UnitsPerPack])"
    GroupLevel =0
    Expression ="CDate(Format(InvoLine.SetupDate,\"mm/dd/yyyy\"))"
    GroupLevel =0
    Expression ="Format(InvoLine.SetupDate,\"mm\")"
    GroupLevel =0
    Expression ="Format(InvoLine.SetupDate,\"yyyy\")"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbMemo "OrderBy" ="qryItemMovementPackSales.PacksPerCarton"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="PacksSold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PackDollars"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MMYY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="YY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PacksPerCarton"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CartonID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PackID"
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
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =144
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =206
        Top =7
        Right =302
        Bottom =121
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =438
        Top =4
        Right =534
        Bottom =118
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
