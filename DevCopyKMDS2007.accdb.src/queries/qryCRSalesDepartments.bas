Operation =1
Option =0
Where ="(((SalesDepartment.ID)>2))"
Begin InputTables
    Name ="SalesDepartment"
End
Begin OutputColumns
    Expression ="SalesDepartment.ID"
    Expression ="SalesDepartment.Description"
End
Begin OrderBy
    Expression ="SalesDepartment.ID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
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
End
