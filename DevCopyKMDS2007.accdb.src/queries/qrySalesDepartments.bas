Operation =1
Option =0
Begin InputTables
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="Item"
    Name ="qryItemsBySalesDepartmentTotals"
End
Begin OutputColumns
    Expression ="SalesGroup.GroupLevel"
    Expression ="SalesGroup.ID"
    Expression ="SalesGroup.Description"
    Expression ="SalesDepartment.ID"
    Expression ="SalesDepartment.Description"
    Expression ="SalesDepartment.Discounts"
    Expression ="SalesDepartment.NegativeOnhand"
    Expression ="SalesDepartment.AllowDiscounts"
    Expression ="SalesDepartment.IncludeInInventoryValueReports"
    Expression ="SalesDepartment.TargetGP"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Alias ="Items"
    Expression ="IIf(IsNull([qryItemsBySalesDepartmentTotals].[count]),0,[qryItemsBySalesDepartme"
        "ntTotals].[count])"
End
Begin Joins
    LeftTable ="SalesGroup"
    RightTable ="SalesDepartment"
    Expression ="SalesGroup.ID = SalesDepartment.SalesGroupID"
    Flag =1
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.ID"
    Flag =2
    LeftTable ="SalesDepartment"
    RightTable ="qryItemsBySalesDepartmentTotals"
    Expression ="SalesDepartment.ID = qryItemsBySalesDepartmentTotals.SalesDepartmentID"
    Flag =2
End
Begin OrderBy
    Expression ="SalesGroup.GroupLevel"
    Flag =0
    Expression ="SalesGroup.ID"
    Flag =0
    Expression ="SalesDepartment.ID"
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
        dbText "Name" ="SalesGroup.GroupLevel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Items"
        dbLong "AggregateType" ="-1"
    End
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
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
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
    Bottom =157
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =72
        Top =10
        Right =169
        Bottom =178
        Top =0
        Name ="SalesGroup"
        Name =""
    End
    Begin
        Left =234
        Top =16
        Right =330
        Bottom =130
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =539
        Top =85
        Right =635
        Bottom =199
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =486
        Top =-2
        Right =630
        Bottom =142
        Top =0
        Name ="qryItemsBySalesDepartmentTotals"
        Name =""
    End
End
