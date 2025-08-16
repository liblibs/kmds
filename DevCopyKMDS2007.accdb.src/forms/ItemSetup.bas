Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DataEntry = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7260
    DatasheetFontHeight =10
    ItemSuffix =27
    Left =8790
    Top =2625
    Right =16635
    Bottom =9105
    AfterInsert ="[Event Procedure]"
    RecSrcDt = Begin
        0x33af10eb0793e240
    End
    Caption ="ItemSetup KMDS Software ©  KMDS Software © "
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
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
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            SpecialEffect =3
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
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
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =6480
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =2
                    ListRows =20
                    ListWidth =3240
                    Left =1680
                    Top =180
                    Width =540
                    Height =255
                    ColumnWidth =2310
                    Name ="SalesDepartmentID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SalesDepartment.ID, SalesDepartment.Description FROM SalesDepartment ORDE"
                        "R BY SalesDepartment.ID; "
                    ColumnWidths ="720;2160"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =165
                            Top =180
                            Width =1455
                            Height =240
                            Name ="SalesDepartmentID_Label"
                            Caption ="Sales Department"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1680
                    Top =540
                    Width =3150
                    ColumnWidth =900
                    TabIndex =2
                    Name ="ID"
                    OnExit ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1380
                            Top =540
                            Width =240
                            Height =240
                            Name ="ID_Label"
                            Caption ="ID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1680
                    Top =1620
                    Width =3150
                    ColumnWidth =2310
                    TabIndex =5
                    Name ="Description"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =750
                            Top =1620
                            Width =870
                            Height =240
                            Name ="Description_Label"
                            Caption ="Description"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    Left =1695
                    Top =1980
                    Width =3180
                    ColumnWidth =2310
                    TabIndex =6
                    Name ="Price"
                    OnExit ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1200
                            Top =1980
                            Width =435
                            Height =240
                            Name ="Price_Label"
                            Caption ="Price"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1680
                    Top =900
                    Width =3120
                    TabIndex =3
                    Name ="txtUPC"
                    OnExit ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =795
                            Top =900
                            Width =825
                            Height =240
                            Name ="Label9"
                            Caption ="UPC Code"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    ColumnCount =2
                    ListRows =20
                    ListWidth =3240
                    Left =2340
                    Top =180
                    Width =2490
                    TabIndex =1
                    BoundColumn =1
                    Name ="SalesDepartmentDesc"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SalesDepartment.Description, SalesDepartment.ID FROM SalesDepartment ORDE"
                        "R BY SalesDepartment.Description; "
                    ColumnWidths ="2160;720"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"

                End
                Begin ComboBox
                    OverlapFlags =85
                    ColumnCount =3
                    ListWidth =2880
                    Left =1680
                    Top =1260
                    Width =1560
                    TabIndex =4
                    Name ="cmbSizeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Size.ID, Size.Description, Size.CaseQty FROM [Size];"
                    ColumnWidths ="0;1440;1440"
                    OnEnter ="[Event Procedure]"
                    DefaultValue ="0"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1245
                            Top =1260
                            Width =375
                            Height =240
                            Name ="Label14"
                            Caption ="Size"
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =67
                    Left =5580
                    Top =5700
                    Width =666
                    Height =606
                    TabIndex =16
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
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"
                    UnicodeAccessKey =67

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =6292
                            Top =5700
                            Width =510
                            Height =525
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label65"
                            Caption ="&Close Form"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =68
                    Left =420
                    Top =5700
                    Width =2160
                    Height =606
                    TabIndex =15
                    Name ="cmdDupe"
                    Caption ="&Duplicate Current Item on Setup Screen"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =68

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =93
                    ColumnCount =2
                    ListRows =20
                    Left =1695
                    Top =2400
                    Width =1920
                    TabIndex =7
                    Name ="PriceGroupID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT PriceGroup.ID, PriceGroup.Description FROM PriceGroup ORDER BY PriceGroup"
                        ".Description; "
                    ColumnWidths ="0;1440"
                    Tag ="TabStopSkip"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =675
                            Top =2400
                            Width =960
                            Height =300
                            Name ="Label16"
                            Caption ="Price Group"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =93
                    ColumnCount =2
                    ListRows =20
                    Left =1695
                    Top =2760
                    Width =1920
                    TabIndex =8
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="VendorID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendor.ID, Vendor.Name FROM Vendor ORDER BY Vendor.Name; "
                    ColumnWidths ="0;2880"
                    Tag ="TabStopSkip"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =1020
                            Top =2760
                            Width =615
                            Height =285
                            Name ="Label7"
                            Caption ="Vendor"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =1695
                    Top =3120
                    Width =780
                    TabIndex =9
                    Name ="Cost"
                    Tag ="TabStopSkip"
                    ControlTipText ="Price paid to acquire goods"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =795
                            Top =3120
                            Width =840
                            Height =285
                            Name ="Label23"
                            Caption ="Cost"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =1695
                    Top =3480
                    Width =780
                    TabIndex =10
                    Name ="Deposit"
                    Tag ="TabStopSkip"
                    ControlTipText ="Bottle and can deposit"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =795
                            Top =3480
                            Width =840
                            Height =285
                            Name ="Label25"
                            Caption ="Deposit"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =1695
                    Top =4920
                    Width =795
                    TabIndex =14
                    Name ="UnitsPerCase"
                    DefaultValue ="1"
                    Tag ="TabStopSkip"
                    ControlTipText ="Used to determine repacks"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =540
                            Top =4920
                            Width =1095
                            Height =285
                            Name ="Label27"
                            Caption ="UnitsPerCase"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =1695
                    Top =3840
                    Width =780
                    TabIndex =11
                    Name ="QtyPack"
                    Tag ="TabStopSkip"
                    ControlTipText ="Used for 6-pack pricing   6 @ $2.99"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =795
                            Top =3840
                            Width =840
                            Height =285
                            Name ="Label45"
                            Caption ="QtyPack"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =1695
                    Top =4200
                    Width =780
                    TabIndex =12
                    Name ="QtyPackPrice"
                    Tag ="TabStopSkip"
                    ControlTipText ="Used for 6-pack pricing   6 @ $2.99"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =795
                            Top =4200
                            Width =840
                            Height =285
                            Name ="Label47"
                            Caption =" PackPrice"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =1695
                    Top =4560
                    Width =1950
                    TabIndex =13
                    Name ="StockNumber"
                    Tag ="TabStopSkip"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =840
                            Top =4560
                            Width =795
                            Height =285
                            Name ="Label49"
                            Caption ="Stock#"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =300
                    Top =2280
                    Width =4860
                    Height =3300
                    Name ="Box17"
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =88
                    Left =5220
                    Top =2430
                    TabIndex =17
                    Name ="chkShowExtended"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    UnicodeAccessKey =120

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5445
                            Top =2400
                            Width =1275
                            Height =570
                            Name ="Label19"
                            Caption ="Show E&xtended Fields"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =3960
                    Top =5700
                    Width =576
                    Height =606
                    TabIndex =18
                    Name ="cmdSave"
                    Caption ="Command20"
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
                        0x0000000000070707000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Save Record"
                    UnicodeAccessKey =83

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =4560
                            Top =5760
                            Width =510
                            Height =525
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label21"
                            Caption ="&Save "
                            FontName ="Trebuchet MS"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "ItemSetup.cls"
