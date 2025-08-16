Operation =1
Option =0
Where ="(((Item.ID) Between Forms!PrintInventoryReport!cboItemIDBegin And Forms!PrintInv"
    "entoryReport!cboItemIDEnd) And ((InvoLine.SetupDate) Between Forms!PrintInventor"
    "yReport!txtDateRangeFrom And Forms!PrintInventoryReport!txtDateRangeTo))"
Begin InputTables
    Name ="BuyDown"
    Name ="Item"
    Name ="InvoLine"
End
Begin OutputColumns
    Expression ="Item.ID"
    Expression ="Item.Description"
    Expression ="InvoLine.Quantity"
    Expression ="InvoLine.Price"
    Expression ="InvoLine.SetupDate"
    Alias ="BusDate"
    Expression ="Format(InvoLine.SetupDate,\"mm/dd/yyyy\")"
End
Begin Joins
    LeftTable ="BuyDown"
    RightTable ="Item"
    Expression ="BuyDown.ItemsID = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="InvoLine"
    Expression ="Item.ID = InvoLine.Code"
    Flag =1
End
Begin OrderBy
    Expression ="Item.ID"
    Flag =0
    Expression ="InvoLine.SetupDate"
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
        dbText "Name" ="BusDate"
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
    Top =15
    ColumnsShown =539
    Begin
        Left =172
        Top =-9
        Right =268
        Bottom =105
        Top =0
        Name ="BuyDown"
        Name =""
    End
    Begin
        Left =396
        Top =0
        Right =492
        Bottom =114
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =539
        Top =2
        Right =635
        Bottom =116
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
