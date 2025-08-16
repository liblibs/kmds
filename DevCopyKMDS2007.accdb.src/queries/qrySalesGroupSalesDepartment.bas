Operation =1
Option =0
Begin InputTables
    Name ="SalesDepartment"
    Name ="SalesGroup"
End
Begin OutputColumns
    Expression ="SalesGroup.ID"
    Expression ="SalesGroup.Description"
    Expression ="SalesDepartment.ID"
    Expression ="SalesDepartment.Description"
    Expression ="SalesDepartment.Discounts"
    Expression ="SalesDepartment.NegativeOnhand"
    Expression ="SalesDepartment.AllowDiscounts"
    Expression ="SalesDepartment.IncludeInInventoryValueReports"
    Expression ="SalesDepartment.TargetGP"
    Expression ="SalesDepartment.AllowClubDiscounts"
    Expression ="SalesDepartment.CatchAllItemID"
    Expression ="SalesDepartment.SetupDate"
    Expression ="SalesDepartment.ChangeDate"
End
Begin Joins
    LeftTable ="SalesDepartment"
    RightTable ="SalesGroup"
    Expression ="SalesDepartment.SalesGroupID = SalesGroup.ID"
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
        dbText "Name" ="SalesGroup.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesGroup.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.Discounts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.NegativeOnhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.AllowDiscounts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.IncludeInInventoryValueReports"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.TargetGP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.AllowClubDiscounts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.CatchAllItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.ChangeDate"
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
    Bottom =446
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="SalesGroup"
        Name =""
    End
End
