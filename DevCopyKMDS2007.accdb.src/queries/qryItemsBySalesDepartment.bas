Operation =1
Option =0
Where ="(((SalesDepartment.ID)>2))"
Begin InputTables
    Name ="SalesDepartment"
    Name ="qryItemsBySalesDepartmentTotals"
End
Begin OutputColumns
    Expression ="SalesDepartment.ID"
    Expression ="SalesDepartment.Description"
    Expression ="qryItemsBySalesDepartmentTotals.Count"
End
Begin Joins
    LeftTable ="SalesDepartment"
    RightTable ="qryItemsBySalesDepartmentTotals"
    Expression ="SalesDepartment.ID=qryItemsBySalesDepartmentTotals.SalesDepartmentID"
    Flag =1
End
Begin OrderBy
    Expression ="SalesDepartment.Description"
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
        dbText "Name" ="SalesDepartment.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDepartment.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemsBySalesDepartmentTotals.count"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =375
    Top =84
    Right =1095
    Bottom =605
    Left =-1
    Top =-1
    Right =688
    Bottom =257
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =258
        Bottom =248
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =303
        Top =56
        Right =572
        Bottom =167
        Top =0
        Name ="qryItemsBySalesDepartmentTotals"
        Name =""
    End
End
