Version =20
VersionRequired =20
Begin Form
    MaxButton = NotDefault
    AutoCenter = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =255
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12840
    DatasheetFontHeight =10
    ItemSuffix =139
    Top =600
    Right =13485
    Bottom =9900
    OnUnload ="[Event Procedure]"
    Filter ="((Lookup_SalesDepartmentID.Description=\"Liquor\"))"
    OrderBy ="Item.ID"
    RecSrcDt = Begin
        0x87ac478ff092e240
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="Item"
    Caption ="Items by Code KMDS Software © "
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin Tab
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =16764057
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =4020
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3555
                    Left =1260
                    Top =1740
                    Width =3840
                    Height =300
                    TabIndex =9
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="SubCategoryID"
                    ControlSource ="SubCategoryID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SubCategory.ID, SubCategory.Description FROM SubCategory ORDER BY [Descri"
                        "ption]; "
                    ColumnWidths ="0;3555"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =420
                            Top =1740
                            Width =780
                            Height =300
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Category_Label"
                            Caption ="Category"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =11460
                    Top =180
                    Width =720
                    Height =300
                    TabIndex =15
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =10800
                            Top =180
                            Width =630
                            Height =285
                            Name ="Label1"
                            Caption ="ID"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    ColumnCount =2
                    ListRows =20
                    Left =6120
                    Top =1020
                    Width =1650
                    Height =300
                    TabIndex =10
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="SizeID"
                    ControlSource ="SizeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Size.ID, Size.Description FROM [Size] ORDER BY Size.Description; "
                    ColumnWidths ="0;1440"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    OnNotInList ="[Event Procedure]"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =5430
                            Top =1020
                            Width =630
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label3"
                            Caption ="SizeID"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    ColumnCount =2
                    ListRows =20
                    Left =1260
                    Top =1380
                    Width =3840
                    Height =300
                    ColumnWidth =2850
                    TabIndex =8
                    BoundColumn =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="SalesDepartmentID"
                    ControlSource ="SalesDepartmentID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SalesDepartment.Description, SalesDepartment.ID FROM SalesDepartment ORDE"
                        "R BY SalesDepartment.Description; "
                    ColumnWidths ="2160;720"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    OnNotInList ="[Event Procedure]"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =300
                            Top =1380
                            Width =900
                            Height =300
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label5"
                            Caption ="Sales Dept."
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    ColumnCount =2
                    ListRows =20
                    Left =8670
                    Top =1020
                    Width =1920
                    Height =300
                    ColumnWidth =1950
                    TabIndex =13
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="VendorID"
                    ControlSource ="VendorID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendor.ID, Vendor.Name FROM Vendor ORDER BY Vendor.Name; "
                    ColumnWidths ="0;2880"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    OnClick ="[Event Procedure]"
                    Tag ="TabStopSkip"
                    OnNotInList ="[Event Procedure]"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =7995
                            Top =1020
                            Width =615
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label7"
                            Caption ="Vendor"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    ColumnCount =2
                    ListRows =20
                    Left =6150
                    Top =1380
                    Width =1635
                    Height =300
                    ColumnWidth =2235
                    TabIndex =11
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="PriceGroupID"
                    ControlSource ="PriceGroupID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT PriceGroup.ID, PriceGroup.Description FROM PriceGroup ORDER BY PriceGroup"
                        ".Description; "
                    ColumnWidths ="0;1440"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    OnClick ="[Event Procedure]"
                    Tag ="TabStopSkip"
                    OnNotInList ="[Event Procedure]"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =5130
                            Top =1380
                            Width =960
                            Height =300
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label9"
                            Caption ="Price Group"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =1230
                    Top =1020
                    Width =3840
                    Height =300
                    ColumnWidth =2895
                    TabIndex =7
                    ForeColor =-2147483640
                    Name ="Description"
                    ControlSource ="Description"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =180
                            Top =1020
                            Width =960
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label11"
                            Caption ="Description"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =1200
                    Top =2220
                    Width =810
                    Height =300
                    TabIndex =16
                    ForeColor =-2147483640
                    Name ="Onhand"
                    ControlSource ="Onhand"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =330
                            Top =2220
                            Width =810
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label13"
                            Caption ="Onhand"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =1200
                    Top =2580
                    Width =810
                    Height =300
                    TabIndex =17
                    ForeColor =-2147483640
                    Name ="OrderAt"
                    ControlSource ="OrderAt"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =375
                            Top =2580
                            Width =765
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label15"
                            Caption ="OrderAt"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =1200
                    Top =2940
                    Width =810
                    Height =300
                    TabIndex =18
                    ForeColor =-2147483640
                    Name ="OrderTo"
                    ControlSource ="OrderTo"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =360
                            Top =2940
                            Width =780
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label17"
                            Caption ="OrderTo"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =3060
                    Top =2220
                    Width =780
                    Height =300
                    TabIndex =19
                    ForeColor =-2147483640
                    Name ="Price"
                    ControlSource ="Price"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2160
                            Top =2220
                            Width =840
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label21"
                            Caption ="Price"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =3060
                    Top =2580
                    Width =780
                    Height =300
                    TabIndex =21
                    ForeColor =-2147483640
                    Name ="Cost"
                    ControlSource ="Cost"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2160
                            Top =2580
                            Width =840
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label23"
                            Caption ="Cost"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =3060
                    Top =2940
                    Width =780
                    Height =300
                    TabIndex =22
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Deposit"
                    ControlSource ="Deposit"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2160
                            Top =2940
                            Width =840
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label25"
                            Caption ="Deposit"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =5955
                    Top =2220
                    Width =795
                    Height =300
                    TabIndex =25
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="UnitsPerCase"
                    ControlSource ="UnitsPerCase"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4800
                            Top =2220
                            Width =1095
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label27"
                            Caption ="UnitsPerCase"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =5955
                    Top =2580
                    Width =795
                    Height =300
                    TabIndex =26
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="UnitsPerOrder"
                    ControlSource ="UnitsPerOrder"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4710
                            Top =2580
                            Width =1185
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label29"
                            Caption ="UnitsPerOrder"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =5955
                    Top =2940
                    Width =795
                    Height =300
                    TabIndex =27
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="UnitsPerPack"
                    ControlSource ="UnitsPerPack"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4800
                            Top =2940
                            Width =1095
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label31"
                            Caption ="UnitsPerPack"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =5400
                    Top =3300
                    Width =1350
                    Height =300
                    TabIndex =28
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="SubPack"
                    ControlSource ="SubPack"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4695
                            Top =3300
                            Width =645
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label33"
                            Caption ="Packs "
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =3045
                    Top =3300
                    Width =780
                    Height =300
                    TabIndex =23
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="QtyPack"
                    ControlSource ="QtyPack"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2145
                            Top =3300
                            Width =840
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label45"
                            Caption ="QtyPack"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =3045
                    Top =3660
                    Width =780
                    Height =300
                    TabIndex =24
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="QtyPackPrice"
                    ControlSource ="QtyPackPrice"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2145
                            Top =3660
                            Width =840
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label47"
                            Caption =" PackPrice"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =8640
                    Top =1380
                    Width =1950
                    Height =274
                    TabIndex =14
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="StockNumber"
                    ControlSource ="StockNumber"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =3
                            Left =7785
                            Top =1380
                            Width =795
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label49"
                            Caption ="Stock#"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Top =960
                    Width =10800
                    Height =1140
                    Name ="Box70"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Top =2100
                    Width =2100
                    Height =1860
                    Name ="Box71"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =2100
                    Top =2100
                    Width =2460
                    Height =1860
                    Name ="Box72"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =4740
                    Top =2100
                    Width =2220
                    Height =1860
                    Name ="Box73"
                End
                Begin Rectangle
                    OverlapFlags =127
                    Left =6960
                    Top =2100
                    Width =3840
                    Height =1860
                    Name ="Box74"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =1
                    OverlapFlags =255
                    TextAlign =2
                    Left =3900
                    Top =2220
                    Width =600
                    Height =300
                    TabIndex =20
                    BackColor =255
                    Name ="txtGP"
                    Format ="Percent"
                    FontName ="Trebuchet MS"

                End
                Begin Subform
                    OverlapFlags =247
                    Left =7020
                    Top =2160
                    Width =2100
                    Height =1740
                    TabIndex =29
                    Name ="UPCItem"
                    SourceObject ="Form.ItemUPC"
                    LinkChildFields ="ItemID"
                    LinkMasterFields ="ID"
                    OnExit ="[Event Procedure]"

                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =93
                    AccessKey =67
                    ColumnCount =7
                    ListRows =25
                    ListWidth =7560
                    Left =780
                    Top =180
                    Width =930
                    Height =285
                    ColumnWidth =1095
                    BackColor =16777088
                    Name ="cmbItemByCode"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.ID, Item.Description, Size.Description AS SizeD, Item.Onhand, Item.P"
                        "rice, Item.Cost, Format(([price]-[cost])/[price],\"Percent\") AS GP FROM Item IN"
                        "NER JOIN [Size] ON Item.SizeID=Size.ID ORDER BY Item.ID; "
                    ColumnWidths ="720;2880;720;720;720;720;360"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =67

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =300
                            Top =180
                            Width =450
                            Height =285
                            ForeColor =16711680
                            Name ="cmbItemByCode_Label"
                            Caption ="&Code"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =95
                    AccessKey =68
                    ColumnCount =7
                    ListRows =25
                    ListWidth =7560
                    Left =2700
                    Top =180
                    Width =2220
                    Height =285
                    ColumnWidth =1785
                    TabIndex =1
                    BoundColumn =2
                    BackColor =16777088
                    Name ="cmbItemByDesc"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.Description, Size.Description AS SizeD, Item.ID, Item.Onhand, Item.P"
                        "rice, Item.Cost, Format(([price]-[cost])/[price],\"Percent\") AS GP FROM Item IN"
                        "NER JOIN [Size] ON Item.SizeID=Size.ID ORDER BY Item.Description; "
                    ColumnWidths ="2880;720;720;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =68

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =1755
                            Top =180
                            Width =945
                            Height =285
                            ForeColor =16711680
                            Name ="Description_Label"
                            Caption ="&Description"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =85
                    ColumnCount =8
                    ListRows =25
                    ListWidth =9000
                    Left =5685
                    Top =180
                    Width =1410
                    Height =285
                    TabIndex =2
                    BoundColumn =3
                    BackColor =16777088
                    Name ="cmbItemByUPC"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UPCItem.ID AS UPCCode, Item.Description, Size.Description AS SizeD, Item."
                        "ID, Item.Onhand, Item.Price, Item.Cost, Format(([price]-[cost])/[price],\"Percen"
                        "t\") AS GP FROM (Item INNER JOIN UPCItem ON Item.ID=UPCItem.ItemID) INNER JOIN ["
                        "Size] ON Item.SizeID=Size.ID ORDER BY UPCItem.ID; "
                    ColumnWidths ="1440;2880;720;720;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =85

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =5295
                            Top =180
                            Width =375
                            Height =285
                            ForeColor =16711680
                            Name ="ID_Label"
                            Caption ="&UPC"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =120
                    Width =10740
                    Height =780
                    Name ="Box78"
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =83
                    ColumnCount =8
                    ListRows =25
                    ListWidth =9720
                    Left =8040
                    Top =180
                    Width =2580
                    Height =285
                    TabIndex =3
                    BoundColumn =3
                    BackColor =16777088
                    Name ="cmbItemBySalesDept"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SalesDepartment.Description AS Dept, Item.Description, Size.Description A"
                        "S SizeD, Item.ID, Item.Onhand, Item.Price, Item.Cost, Format(([price]-[cost])/[p"
                        "rice],\"Percent\") AS GP FROM (SalesDepartment INNER JOIN Item ON SalesDepartmen"
                        "t.ID=Item.SalesDepartmentID) INNER JOIN [Size] ON Item.SizeID=Size.ID ORDER BY S"
                        "alesDepartment.Description, Item.Description; "
                    ColumnWidths ="1440;2880;720;720;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =83

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =7140
                            Top =180
                            Width =840
                            Height =285
                            ForeColor =16711680
                            Name ="Label81"
                            Caption ="&Sales Dept"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =80
                    ColumnCount =8
                    ListRows =25
                    ListWidth =9000
                    Left =2715
                    Top =480
                    Width =2205
                    Height =285
                    TabIndex =4
                    BoundColumn =3
                    BackColor =16777088
                    Name ="cmbItemsByPriceGroup"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT PriceGroup.Description AS PriceGroup, Item.Description, Size.Description "
                        "AS SizeD, Item.ID, Item.Onhand, Item.Price, Item.Cost, Format(([price]-[cost])/["
                        "price],\"Percent\") AS GP FROM (Item INNER JOIN PriceGroup ON Item.PriceGroupID="
                        "PriceGroup.ID) INNER JOIN [Size] ON Item.SizeID=Size.ID ORDER BY PriceGroup.Desc"
                        "ription, Item.Description; "
                    ColumnWidths ="1440;2880;720;720;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =80

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =1695
                            Top =480
                            Width =1005
                            Height =285
                            ForeColor =16711680
                            Name ="Label83"
                            Caption =" &Price Group"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =76
                    ColumnCount =8
                    ListRows =25
                    ListWidth =8280
                    Left =5700
                    Top =480
                    Width =1395
                    Height =285
                    TabIndex =5
                    BoundColumn =3
                    BackColor =16777088
                    Name ="cmbItemByLocation"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.Location, Item.Description, Size.Description AS SizeD, Item.ID, Item"
                        ".Onhand, Item.Price, Item.Cost, Format(([price]-[cost])/[price],\"Percent\") AS "
                        "GP FROM Item INNER JOIN [Size] ON Item.SizeID=Size.ID ORDER BY Item.Location, It"
                        "em.Description; "
                    ColumnWidths ="720;2880;720;720;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =76

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =4920
                            Top =480
                            Width =735
                            Height =285
                            ForeColor =16711680
                            Name ="Label87"
                            Caption ="&Location"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =89
                    ColumnCount =8
                    ListRows =25
                    ListWidth =9360
                    Left =8025
                    Top =480
                    Width =2625
                    Height =285
                    TabIndex =6
                    BoundColumn =3
                    BackColor =16777088
                    Name ="cmbItemBySubCategoryID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SubCategory.Description AS Category, Item.Description, Size.Description A"
                        "S SizeD, Item.ID, Item.Onhand, Item.Price, Item.Cost, Format(([price]-[cost])/[p"
                        "rice],\"Percent\") AS GP FROM (Item INNER JOIN SubCategory ON Item.SubCategoryID"
                        "=SubCategory.ID) INNER JOIN [Size] ON Item.SizeID=Size.ID ORDER BY SubCategory.D"
                        "escription, Item.Description, Item.ID; "
                    ColumnWidths ="2160;2880;720;720;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =121

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =7200
                            Top =480
                            Width =780
                            Height =285
                            ForeColor =16711680
                            Name ="Label101"
                            Caption ="Categor&y"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin OptionButton
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =10200
                    Top =1740
                    TabIndex =30
                    Name ="optTabSkip"
                    DefaultValue ="0"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =8100
                            Top =1740
                            Width =2025
                            Height =240
                            Name ="Label105"
                            Caption ="Stop on Specified Fields"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2880
                    Left =1200
                    Top =3300
                    Width =840
                    TabIndex =31
                    BackColor =255
                    Name ="OnOrder"
                    ControlSource ="ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.ID, Sum(IIf(IsNull([AgentTransferLines].[QuantityShipped]),0,[AgentT"
                        "ransferLines].[QuantityShipped]))+Sum(IIf(IsNull([PurchaseOrderLine].[QuantitySh"
                        "ipped]),0,[PurchaseOrderLine].[QuantityShipped])) AS OnOrder, Sum(IIf(IsNull([Ag"
                        "entTransferLines].[QuantityShipped]),0,[AgentTransferLines].[QuantityShipped])) "
                        "AS OnXFR, Sum(IIf(IsNull([PurchaseOrderLine].[QuantityShipped]),0,[PurchaseOrder"
                        "Line].[QuantityShipped])) AS OO FROM (Item LEFT JOIN AgentTransferLines ON Item."
                        "ID=AgentTransferLines.ItemID) LEFT JOIN PurchaseOrderLine ON Item.ID=PurchaseOrd"
                        "erLine.ItemID GROUP BY Item.ID; "
                    ColumnWidths ="0;1440;1440"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ControlTipText ="Double click on this number to see what Purchase Order(s) have this item."

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =390
                            Top =3300
                            Width =780
                            Height =245
                            Name ="On Order_Label"
                            Caption ="On Order"
                            FontName ="Trebuchet MS"
                            EventProcPrefix ="On_Order_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2880
                    Left =1200
                    Top =3665
                    Width =840
                    TabIndex =32
                    BackColor =255
                    Name ="OnReserve"
                    ControlSource ="ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT OrderLine.ItemID, Sum(OrderLine.QuantityShipped) AS SumOfQuantityShipped "
                        "FROM OrderLine GROUP BY OrderLine.ItemID; "
                    ColumnWidths ="0;1440;1440"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ControlTipText ="Double click on this number to see what Order(s) have this item."

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =240
                            Top =3660
                            Width =930
                            Height =255
                            Name ="Label110"
                            Caption ="On Reserve"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =69
                    Left =10860
                    Top =2580
                    Width =666
                    Height =606
                    TabIndex =33
                    ForeColor =0
                    Name ="cmdfrmDeleteRecord"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada177adada77da1dad1177adad17ad11da7117dad71ada ,
                        0x111da1177d117dad1111d7117711dada11111d11111dadad1111da71117adada ,
                        0x111d77111177adad11d711da71177ada1dadadada71177addadadadadad11ada ,
                        0xadadadadadadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"
                    UnicodeAccessKey =101

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =11520
                            Top =2580
                            Width =660
                            Height =495
                            Name ="lblDeleteRecord"
                            Caption ="D&elete Item"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =86
                    Left =10860
                    Top =1920
                    Width =666
                    Height =606
                    TabIndex =34
                    ForeColor =0
                    Name ="cmdfrmSaveRecord"
                    Caption ="Command34"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f8f8070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f8f8f8070700000000000000000000000000000000000000000000 ,
                        0x0000000000f8f8f8070007070707070700f8f8f8f8f8f8f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800000000f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800000000f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f8f8f8f8f8f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700000000000000000000000000000000 ,
                        0x000007070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070000000000000000000000000000000000000000 ,
                        0x000007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007000700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007000700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8070007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff00070707000707070000000000000000000000000000000000000000000000 ,
                        0x0000000000070707
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Save Changes to Record"
                    UnicodeAccessKey =118

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =11535
                            Top =2100
                            Width =555
                            Height =210
                            Name ="lblSaveRecord"
                            Caption ="Sa&ve"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =78
                    Left =10860
                    Top =1260
                    Width =666
                    Height =606
                    TabIndex =35
                    ForeColor =0
                    Name ="cmdfrmUndo"
                    Caption ="Command35"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070700 ,
                        0xfb01fb01000707070707070707070707070707070707070707070707070700fb ,
                        0x01fb01000107070707070707070707070707070707070707070707070700fb01 ,
                        0xfb010001fb070707070707070707070707070707070707070707070700fb01fb ,
                        0x010001fb010707070707070707070707070707070707070707070700fb01fb01 ,
                        0x0001fb01fb07070707070707070707070707070707070707070700fb01fb0100 ,
                        0x01fb01fb00070707070707070707070707070707070707070700fb01fb010001 ,
                        0xfb01fb00fb0707070707070707070707070707070707070700fb01fb010001fb ,
                        0x01fb00fb0007070707070707070707070707070707070700000000010001fb01 ,
                        0xfb00fb0007070707070707070707070707070707070700020202000001fb01fb ,
                        0x00fb00070707070707070707070707070707070707000000000202020001fb00 ,
                        0xfb0007070707070707070707070707070707070700fb01fb00000202000000fb ,
                        0x0007070707070707070707070707070707070700fb01fb01fb0100020200fb00 ,
                        0x07070707070707070707070707070707070700fb01fb01fb01fb000002000007 ,
                        0x07070707070707070707070707070707070000000001fb01fb01fb0002000707 ,
                        0x0707070707070707070707070707070700020202000001fb01fb010000070707 ,
                        0x070707070707070707070707070707000000000202020001fb01fb0007070707 ,
                        0x07070707070707070707070707070007fd0700000202000001fb000707070707 ,
                        0x070707070707070707070707070007fd07fd07fd00020200fb00070707070707 ,
                        0x070707070707070707070707000000000007fd07000002000007070707070707 ,
                        0x07070707070707070707070700fdfffdff0007fd070002000707070707070707 ,
                        0x070707070707070707070700fdfffdfffd00fd07fd0000070707070707070707 ,
                        0x070707070707070707070700fffdfffdff0000fd070007070707070707070707 ,
                        0x070707070707070707070700fdfffdfffdff0007000707070707070707070707 ,
                        0x070707070707070707070700fffdfffdfffd0000070707070707070707070707 ,
                        0x07070707070707070707070700fffdfffdff0007070707070707070707070707 ,
                        0x07070707070707070707070700fdfffdff000707070707070707070707070707 ,
                        0x0707070707070707070707070700000000070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707
                    End
                    FontName ="Trebuchet MS"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Undo Changes Made to This Record"
                    UnicodeAccessKey =110

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =11535
                            Top =1380
                            Width =585
                            Height =255
                            Name ="lblUndo"
                            Caption ="U&ndo"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Top =120
                    Width =12840
                    Height =3840
                    Name ="boxScreen"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    AccessKey =77
                    Left =10860
                    Top =3300
                    Width =666
                    Height =606
                    TabIndex =36
                    ForeColor =0
                    Name ="cmdfrmClose"
                    Caption ="Command64"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f907 ,
                        0x07070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0x070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf907070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f90707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f907070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f90707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f90707f9f9f9f9fffffffff9f9f9fffff9f9f9fffffffff9f9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9f9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9f9fffffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9fffffff9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfffffffff9f9f90707f9f9f9f9fffffff9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffffff9f9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9f9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9f9fffffffff9fffffffffffff9fffffffff9f9ff ,
                        0xfffffffff9f9f90707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f9070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f90707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f907070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f90707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f907070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0x0707070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f907 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"
                    UnicodeAccessKey =109

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =11572
                            Top =3300
                            Width =510
                            Height =525
                            Name ="Label65"
                            Caption ="Close For&m"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    AccessKey =65
                    Left =10860
                    Top =600
                    Width =666
                    Height =606
                    TabIndex =37
                    ForeColor =0
                    Name ="cmdfrmAddNew"
                    Caption ="Command67"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddcdadadadadadadaaccdadadada4adaddcccdadad4d4d4da ,
                        0xaccccdadad444daddcccccda4444444aaccccdadad444daddcccdadad4d4d4da ,
                        0xaccdadadada4adaddcdadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Add a New Record"
                    UnicodeAccessKey =65

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =11543
                            Top =780
                            Width =705
                            Height =285
                            Name ="Label68"
                            Caption ="&Add New"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    AccessKey =82
                    ColumnCount =8
                    ListRows =25
                    ListWidth =9360
                    Left =780
                    Top =480
                    Width =930
                    Height =285
                    TabIndex =38
                    BoundColumn =3
                    BackColor =16776960
                    Name ="cmbItemsByVendor"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendor.Name, Item.Description, Size.Description AS SizeD, Item.ID, Item.O"
                        "nhand, Item.Price, Item.Cost, Format(([price]-[cost])/[price],\"Percent\") AS GP"
                        " FROM (Item INNER JOIN Vendor ON Item.VendorID=Vendor.ID) INNER JOIN [Size] ON I"
                        "tem.SizeID=Size.ID ORDER BY Vendor.Name, Item.Description; "
                    ColumnWidths ="1440;2880;720;720;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =114

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =120
                            Top =480
                            Width =615
                            Height =285
                            ForeColor =16711680
                            Name ="Label114"
                            Caption ="Vendo&r"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =247
                    Left =6120
                    Top =1740
                    Width =810
                    Height =300
                    TabIndex =12
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Location"
                    ControlSource ="Location"
                    OnDblClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    Tag ="TabStopSkip"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =5250
                            Top =1740
                            Width =810
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label43"
                            Caption ="Location"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =10170
                    Top =2220
                    Width =510
                    Height =300
                    TabIndex =39
                    BackColor =8454016
                    ForeColor =-2147483640
                    Name ="txtAvailOnhand"
                    ControlSource ="Onhand"
                    FontName ="MS Sans Serif"
                    Tag =" "
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =9300
                            Top =2220
                            Width =810
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label123"
                            Caption ="Onhand"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =10185
                    Top =2640
                    Width =510
                    Height =300
                    TabIndex =40
                    BackColor =8454016
                    Name ="txtAvailSoldToday"
                    ControlSource ="=[SoldTodayCr1]+[SoldTodayCr2]+[SoldTodayCr3]+[SoldTodayCr4]"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =9255
                            Top =2640
                            Width =855
                            Height =285
                            Name ="Label125"
                            Caption ="Sold Today"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =10185
                    Top =3000
                    Width =510
                    Height =300
                    TabIndex =41
                    BackColor =8454016
                    Name ="txtAvailTransferred"
                    ControlSource ="TransferredToday"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =9255
                            Top =3000
                            Width =900
                            Height =285
                            Name ="Label131"
                            Caption ="XFR Today"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =10185
                    Top =3420
                    Width =510
                    Height =300
                    TabIndex =42
                    BackColor =8454016
                    Name ="txtAvailOnTheShelf"
                    ControlSource ="=[onhand]+[TransferredToday]-[SoldTodayCr1]-[SoldTodayCr2]-[SoldTodayCr3]-[SoldT"
                        "odayCr4]"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =9180
                            Top =3420
                            Width =975
                            Height =285
                            Name ="Label133"
                            Caption ="On the shelf"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =6045
                    Top =3660
                    Width =250
                    TabIndex =43
                    Name ="PLUList"
                    ControlSource ="PLUList"
                    Tag ="TabStopSkip"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =5220
                            Top =3660
                            Width =765
                            Height =285
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label51"
                            Caption ="PLU List"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =2760
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =3120
                    Top =60
                    Width =1740
                    Height =300
                    TabIndex =1
                    BackColor =255
                    Name ="DateLastSold"
                    ControlSource ="DateLastSold"
                    Format ="mm/dd/yy hh:nn:ss"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2340
                            Top =60
                            Width =720
                            Height =285
                            Name ="Label53"
                            Caption ="LastSold"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    Left =600
                    Top =60
                    Width =1740
                    Height =300
                    ColumnWidth =3090
                    BackColor =255
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    Format ="mm/dd/yy hh:nn:ss"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =60
                            Width =495
                            Height =285
                            Name ="Label55"
                            Caption ="Setup"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =9480
                    Top =60
                    Width =1740
                    Height =300
                    ColumnWidth =2070
                    TabIndex =2
                    BackColor =255
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    Format ="mm/dd/yy hh:nn:ss"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8700
                            Top =60
                            Width =720
                            Height =285
                            Name ="Label57"
                            Caption ="Changed"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =127
                    TextAlign =2
                    Left =5700
                    Top =60
                    Width =360
                    Height =300
                    TabIndex =3
                    BackColor =255
                    Name ="SoldTodayCr1"
                    ControlSource ="SoldTodayCr1"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =4800
                            Top =60
                            Width =900
                            Height =285
                            Name ="Label35"
                            Caption ="Sold Today"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =127
                    TextAlign =2
                    Left =6060
                    Top =60
                    Width =360
                    Height =300
                    TabIndex =4
                    BackColor =255
                    Name ="SoldTodayCr2"
                    ControlSource ="SoldTodayCr2"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =127
                    TextAlign =2
                    Left =6420
                    Top =60
                    Width =360
                    Height =300
                    TabIndex =5
                    BackColor =255
                    Name ="SoldTodayCr3"
                    ControlSource ="SoldTodayCr3"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =119
                    TextAlign =2
                    Left =6780
                    Top =60
                    Width =360
                    Height =300
                    TabIndex =6
                    BackColor =255
                    Name ="SoldTodayCr4"
                    ControlSource ="SoldTodayCr4"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =8220
                    Top =60
                    Width =360
                    Height =300
                    TabIndex =7
                    BackColor =255
                    Name ="TransferredToday"
                    ControlSource ="TransferredToday"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7260
                            Top =60
                            Width =900
                            Height =285
                            Name ="Label117"
                            Caption ="XFR Today"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =12660
                    Top =1380
                    Width =180
                    Height =615
                    FontSize =12
                    BackColor =16763904
                    Name ="lblSaving"
                    Caption ="Saving Record"
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =12660
                    Top =540
                    Width =180
                    Height =615
                    FontSize =12
                    BackColor =16763904
                    Name ="lblEditing"
                    Caption ="Editing Record"
                End
                Begin Subform
                    OverlapFlags =85
                    Left =60
                    Top =720
                    Width =12480
                    Height =2040
                    TabIndex =8
                    Name ="qryItemMovementSalesTotal"
                    SourceObject ="Form.ItemChartSalesTotal"
                    LinkChildFields ="Code"
                    LinkMasterFields ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =420
                            Width =2940
                            Height =240
                            Name ="Label137"
                            Caption ="Sales total by month"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11340
                    Top =120
                    Width =1260
                    Height =525
                    TabIndex =9
                    Name ="cmdItemForm"
                    Caption ="Show Item Movement"
                    OnClick ="[Event Procedure]"

                End
            End
        End
    End
End
CodeBehindForm
' See "ItemWithSalesTotalChart.cls"
