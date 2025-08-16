Operation =1
Option =0
Where ="(((Item.RestockAmount)<>0) AND ((SalesDepartment.IncludeInInventoryValueReports)"
    "<>0))"
Begin InputTables
    Name ="Item"
    Name ="qryItemOnReserve"
    Name ="SalesDepartment"
End
Begin OutputColumns
    Alias ="ID"
    Expression ="IIf([SalesdepartmentID]=1,([item].[id]-99900000000),[item].[id])"
    Expression ="Item.Description"
    Alias ="OH"
    Expression ="[Onhand]-[SoldTodayCr1]-[SoldTodayCr2]-[SoldTodayCr3]-[SoldTodayCr4]+[Transferre"
        "dToday]"
    Alias ="Reserved"
    Expression ="IIf(IsNull([OnReserve]),0,[OnReserve])"
    Expression ="Item.RestockAmount"
    Expression ="SalesDepartment.IncludeInInventoryValueReports"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="qryItemOnReserve"
    Expression ="Item.ID = qryItemOnReserve.ItemID"
    Flag =2
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Item.SalesDepartmentID"
    Flag =0
    Expression ="Item.ID"
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
        dbText "Name" ="OH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reserved"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.RestockAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.IncludeInInventoryValueReports"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID"
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
    Right =1193
    Bottom =260
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =259
        Bottom =300
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =296
        Top =6
        Right =515
        Bottom =90
        Top =0
        Name ="qryItemOnReserve"
        Name =""
    End
    Begin
        Left =305
        Top =106
        Right =497
        Bottom =295
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
End
