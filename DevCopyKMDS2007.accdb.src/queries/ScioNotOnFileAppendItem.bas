Operation =3
Name ="Item"
Option =0
Begin InputTables
    Name ="ScioNotOnFile"
End
Begin OutputColumns
    Name ="ID"
    Expression ="ScioNotOnFile.ID"
    Name ="SizeID"
    Expression ="ScioNotOnFile.SizeID"
    Name ="VendorID"
    Expression ="ScioNotOnFile.VendorID"
    Name ="Description"
    Expression ="ScioNotOnFile.Description"
    Name ="Price"
    Expression ="ScioNotOnFile.Price"
    Name ="Cost"
    Expression ="ScioNotOnFile.Cost"
    Name ="UnitsPerCase"
    Expression ="ScioNotOnFile.UnitsPerCase"
    Name ="UnitsPerOrder"
    Expression ="ScioNotOnFile.UnitsPerOrder"
    Name ="SalesDepartmentId"
    Expression ="ScioNotOnFile.SalesDepartmentId"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ScioNotOnFile.UnitsPerOrder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScioNotOnFile.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScioNotOnFile.SizeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScioNotOnFile.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScioNotOnFile.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScioNotOnFile.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScioNotOnFile.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScioNotOnFile.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScioNotOnFile.SalesDepartmentId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1531
    Bottom =873
    Left =-1
    Top =-1
    Right =1508
    Bottom =563
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =103
        Top =140
        Right =247
        Bottom =284
        Top =0
        Name ="ScioNotOnFile"
        Name =""
    End
End
