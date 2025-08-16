Operation =1
Option =0
Where ="(((OrdHead.InvoiceNo)=Forms!ClubOrderPrintDialog!txtInvoiceNo) And ((OrderLine.Q"
    "uantityShipped)<>0) And ((Item.SalesDepartmentID)=1))"
Begin InputTables
    Name ="OrdHead"
    Name ="Item"
    Name ="Size"
    Name ="OrderLine"
End
Begin OutputColumns
    Expression ="Size.Description"
    Alias ="Count"
    Expression ="Sum(1)"
    Alias ="CQ"
    Expression ="First(Size.CaseQty)"
    Alias ="QS"
    Expression ="Sum(OrderLine.QuantityShipped)"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =2
    LeftTable ="Item"
    RightTable ="OrderLine"
    Expression ="Item.ID = OrderLine.ItemID"
    Flag =1
    LeftTable ="OrdHead"
    RightTable ="OrderLine"
    Expression ="OrdHead.ID = OrderLine.OrdheadID"
    Flag =1
End
Begin Groups
    Expression ="Size.Description"
    GroupLevel =0
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
        dbText "Name" ="Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CQ"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QS"
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
        Left =24
        Top =2
        Right =120
        Bottom =116
        Top =0
        Name ="OrdHead"
        Name =""
    End
    Begin
        Left =379
        Top =2
        Right =475
        Bottom =116
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =489
        Top =-5
        Right =585
        Bottom =109
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =192
        Top =-7
        Right =288
        Bottom =107
        Top =0
        Name ="OrderLine"
        Name =""
    End
End
