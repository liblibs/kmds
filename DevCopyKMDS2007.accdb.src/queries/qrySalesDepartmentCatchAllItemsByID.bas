Operation =1
Option =0
Where ="(((SalesDepartment.ID)=[parSalesDepartmentID]))"
Begin InputTables
    Name ="SalesDepartment"
    Name ="Item"
End
Begin OutputColumns
    Expression ="SalesDepartment.ID"
    Expression ="SalesDepartment.CatchAllItemID"
End
Begin Parameters
    Name ="parSalesDepartmentID"
    Flag =4
End
Begin Joins
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.CatchAllItemID = Item.ID"
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
        dbText "Name" ="SalesDepartment.ID"
        dbInteger "ColumnWidth" ="390"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.CatchAllItemID"
        dbInteger "ColumnWidth" ="2580"
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
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =161
        Top =6
        Right =257
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
End
