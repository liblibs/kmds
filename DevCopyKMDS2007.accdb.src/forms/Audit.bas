Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =48
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15300
    DatasheetFontHeight =10
    ItemSuffix =83
    Left =2220
    Top =1020
    Right =17115
    Bottom =10215
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x63c1ea9d4fdee240
    End
    Caption =" KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Softwa"
        "re ©  KMDS Software ©  KMDS Software © "
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
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
        Begin FormHeader
            Height =2070
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =60
                    Width =3300
                    Height =480
                    FontSize =16
                    Name ="Label15"
                    Caption ="Inventory Adjustment "
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =4485
                    Top =420
                    Width =1875
                    ColumnOrder =0
                    BackColor =10092543
                    Name ="cmbListType"
                    RowSourceType ="Value List"
                    RowSource ="Code;Description;SalesDept;Vendor;PriceGroup;Location;Category;Size"
                    ColumnWidths ="1440"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3600
                            Top =420
                            Width =840
                            Height =240
                            Name ="Label17"
                            Caption ="Range of:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =25
                    ListWidth =1440
                    Left =7320
                    Top =420
                    Width =1860
                    ColumnOrder =1
                    TabIndex =1
                    Name ="cmbSubCategoryID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SalesDepartment.Description, SalesDepartment.ID, SalesDepartment.ID FROM "
                        "SalesDepartment ORDER BY SalesDepartment.Description; "
                    ColumnWidths ="1440;720;0"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6420
                            Top =420
                            Width =840
                            Height =240
                            Name ="Description_Label"
                            Caption ="SalesDept"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =25
                    ListWidth =1440
                    Left =7335
                    Top =720
                    Width =1860
                    ColumnOrder =2
                    TabIndex =2
                    Name ="cmbSubCategoryIDEnd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SalesDepartment.Description, SalesDepartment.ID, SalesDepartment.ID FROM "
                        "SalesDepartment ORDER BY SalesDepartment.Description; "
                    ColumnWidths ="1440;720;0"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7020
                            Top =720
                            Width =255
                            Height =240
                            Name ="Label20"
                            Caption ="To"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =9300
                    Top =420
                    Width =1380
                    Height =780
                    TabIndex =3
                    Name ="cmdAddItems"
                    Caption ="Add items from this range of SalesDepts."
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Run Query"

                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =60
                    Top =780
                    Width =4322
                    Height =361
                    ColumnOrder =3
                    TabIndex =4
                    Name ="framSortBy"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            Left =180
                            Top =660
                            Width =720
                            Height =240
                            BackColor =-2147483633
                            Name ="Label25"
                            Caption ="Sort by"
                        End
                        Begin OptionButton
                            OverlapFlags =215
                            Left =300
                            Top =868
                            OptionValue =1
                            Name ="Option27"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =525
                                    Top =840
                                    Width =615
                                    Height =240
                                    Name ="Label28"
                                    Caption ="Item ID"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =1260
                            Top =868
                            OptionValue =2
                            Name ="Option29"
                            OnGotFocus ="[Event Procedure]"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1490
                                    Top =840
                                    Width =870
                                    Height =240
                                    Name ="Label30"
                                    Caption ="Description"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =2460
                            Top =868
                            OptionValue =3
                            Name ="Option31"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =2690
                                    Top =840
                                    Width =360
                                    Height =240
                                    Name ="Label32"
                                    Caption ="Size"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =3180
                            Top =868
                            OptionValue =4
                            Name ="Option50"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =3410
                                    Top =840
                                    Width =885
                                    Height =240
                                    Name ="Label51"
                                    Caption ="Setup Date"
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =3660
                    Top =120
                    Width =7260
                    Height =210
                    Name ="Label33"
                    Caption ="Fill the list of items by selecting a range of items with the options below..."
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =12840
                    Top =1830
                    ColumnOrder =4
                    TabIndex =5
                    Name ="chkSkipPhysical"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =13065
                            Top =1800
                            Width =1155
                            Height =240
                            Name ="Label42"
                            Caption ="Skip Physical "
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =12840
                    Top =1590
                    ColumnOrder =5
                    TabIndex =6
                    Name ="chkSkipVariance"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =13065
                            Top =1560
                            Width =1020
                            Height =240
                            Name ="Label45"
                            Caption ="Skip Variance"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =180
                    Top =1380
                    Width =2940
                    Height =630
                    ColumnOrder =6
                    FontSize =24
                    TabIndex =7
                    BackColor =8454143
                    BorderColor =8388863
                    Name ="txtDispItem"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3240
                    Top =1380
                    Width =7680
                    Height =630
                    ColumnOrder =7
                    FontSize =24
                    TabIndex =8
                    BackColor =8454143
                    BorderColor =8388863
                    Name ="txtDispDescription"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =1380
                    Width =1680
                    Height =660
                    ColumnOrder =8
                    FontSize =24
                    TabIndex =9
                    BackColor =8454143
                    BorderColor =8388863
                    Name ="txtDispOnShelf"

                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =11220
                    Top =180
                    Width =1442
                    Height =1138
                    ColumnOrder =9
                    TabIndex =10
                    Name ="optCountBy"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =11340
                            Top =60
                            Width =1215
                            Height =240
                            BackColor =-2147483633
                            Name ="Label61"
                            Caption ="Enter counts by"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =11286
                            Top =448
                            OptionValue =1
                            Name ="Option63"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =11516
                                    Top =420
                                    Width =1125
                                    Height =240
                                    Name ="Label64"
                                    Caption ="Physical Count"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =11280
                            Top =718
                            OptionValue =2
                            Name ="Option65"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =11510
                                    Top =690
                                    Width =690
                                    Height =240
                                    Name ="Label66"
                                    Caption ="Variance"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =11280
                            Top =1018
                            OptionValue =3
                            Name ="Option67"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =11510
                                    Top =990
                                    Width =750
                                    Height =240
                                    Name ="Label68"
                                    Caption ="Locations"
                                End
                            End
                        End
                    End
                End
                Begin OptionGroup
                    Visible = NotDefault
                    OverlapFlags =119
                    Left =12840
                    Top =180
                    Width =1322
                    Height =1378
                    ColumnOrder =10
                    TabIndex =11
                    Name ="optLocations"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            Left =12960
                            Top =60
                            Width =675
                            Height =240
                            BackColor =-2147483633
                            Name ="Label70"
                            Caption ="Location"
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =13026
                            Top =418
                            OptionValue =1
                            Name ="Option72"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =13256
                                    Top =390
                                    Width =465
                                    Height =240
                                    Name ="Label73"
                                    Caption ="Front"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =13020
                            Top =688
                            OptionValue =2
                            Name ="Option74"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =13250
                                    Top =660
                                    Width =405
                                    Height =240
                                    Name ="Label75"
                                    Caption ="Back"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =13020
                            Top =988
                            OptionValue =3
                            Name ="Option76"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =13250
                                    Top =960
                                    Width =585
                                    Height =240
                                    Name ="Label77"
                                    Caption ="Display"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =13020
                            Top =1288
                            OptionValue =4
                            Name ="Option78"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =13250
                                    Top =1260
                                    Width =735
                                    Height =240
                                    Name ="Label79"
                                    Caption ="Overflow"
                                End
                            End
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5640
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =215
                    Left =60
                    Top =120
                    Width =14280
                    Height =5400
                    Name ="AuditSub"
                    SourceObject ="Form.AuditSub"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =60
                            Top =480
                            Width =720
                            Height =240
                            Name ="Label0"
                            Caption ="AuditSub"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =223
                    Left =60
                    Top =60
                    Width =14400
                    Height =5520
                    Name ="Box80"
                End
            End
        End
        Begin FormFooter
            Height =1500
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1920
                    Top =360
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Text1"
                    ControlSource ="=AuditSub.Form!TotOnShelf"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1920
                            Top =60
                            Width =630
                            Height =240
                            Name ="Label2"
                            Caption ="Onshelf"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3240
                    Top =360
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Text3"
                    ControlSource ="=AuditSub.Form!TotPhysicalCount"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3240
                            Top =60
                            Width =645
                            Height =240
                            Name ="Label4"
                            Caption ="Physical"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7200
                    Top =360
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Text5"
                    ControlSource ="=AuditSub.Form!TotExtCost"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000000 ,
                        0x00000000ff99cc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000000000000050000000100000000000000ff99cc00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7200
                            Top =60
                            Width =630
                            Height =240
                            Name ="Label6"
                            Caption ="Cost"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4560
                    Top =360
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="Text7"
                    ControlSource ="=AuditSub.Form!TotVarianceAmount"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001000000 ,
                        0x00000000ff99cc00000000000400000003000000050000000100000000000000 ,
                        0xccffcc0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000200000000000000050000000100000000000000ff99cc00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000004000000010000 ,
                        0x0000000000ccffcc000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4560
                            Top =60
                            Width =690
                            Height =240
                            Name ="Label8"
                            Caption ="Variance"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5880
                    Top =360
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="Text9"
                    ControlSource ="=AuditSub.Form!TotExtPrice"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001000000 ,
                        0x00000000ff99cc00000000000400000003000000050000000100000000000000 ,
                        0xccffcc0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000200000000000000050000000100000000000000ff99cc00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000004000000010000 ,
                        0x0000000000ccffcc000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5880
                            Top =60
                            Width =870
                            Height =240
                            Name ="Label10"
                            Caption ="Retail Price"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11820
                    Top =300
                    Width =975
                    Height =825
                    TabIndex =5
                    Name ="cmdPost"
                    Caption ="POST This Batch"
                    OnClick ="[Event Procedure]"

                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =77
                    Left =13680
                    Top =120
                    Width =666
                    Height =606
                    TabIndex =6
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
                            OverlapFlags =85
                            TextAlign =2
                            Left =13680
                            Top =720
                            Width =885
                            Height =210
                            Name ="Label65"
                            Caption ="Close For&m"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    Left =120
                    Top =60
                    Width =576
                    Height =576
                    TabIndex =7
                    Name ="Command21"
                    Caption ="cmdDeleteRecords"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada177adada77da1dad1177adad17ad11da7117dad71ada ,
                        0x111da1177d117dad1111d7117711dada11111d11111dadad1111da71117adada ,
                        0x111d77111177adad11d711da71177ada1dadadada71177addadadadadad11ada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"
                    UnicodeAccessKey =68

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =720
                            Top =60
                            Width =885
                            Height =435
                            Name ="Label22"
                            Caption ="&Delete ALL records"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =86
                    Left =9840
                    Top =60
                    Width =576
                    Height =576
                    TabIndex =8
                    Name ="cmdPreview"
                    Caption ="Command34"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada000000000000ad000ffffffffff0d000 ,
                        0x0fffffff0000700d0ffffff0788707da0fffff0788e770ad0fffff08888780da ,
                        0x0fffff08e88780ad0fffff07ee8770da0ffffff078870dad0fffffff0000dada ,
                        0x0ffffffffff0adad0fffffff0000dada0fffffff080dadad0fffffff00dadada ,
                        0x000000000dadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Preview Report"
                    UnicodeAccessKey =118

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =10500
                            Top =120
                            Width =645
                            Height =240
                            Name ="Label37"
                            Caption ="Pre&view"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =9180
                    Top =900
                    Width =245
                    TabIndex =9
                    Name ="chkOnlyVariance"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextAlign =2
                            Left =8820
                            Top =420
                            Width =900
                            Height =480
                            Name ="Label40"
                            Caption ="Only Print Variance"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7260
                    Top =720
                    Width =900
                    Height =720
                    TabIndex =10
                    Name ="cmdAuditUpdateOnShelf"
                    Caption ="Refresh OnShelf "
                    OnClick ="[Event Procedure]"

                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =80
                    Left =9840
                    Top =720
                    Width =576
                    Height =576
                    TabIndex =11
                    Name ="cmdPrint"
                    Caption ="Command34"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaad00000000000dadd0888888888080da ,
                        0x000000000000080d0888888bbb88000a088888877788080d0000000000000880 ,
                        0x0888888888808080d000000000080800ad0ffffffff08080dad0f00000f0000a ,
                        0xada0ffffffff0daddada0f00000f0adaadad0ffffffff0addadad000000000da ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Preview Report"
                    UnicodeAccessKey =80

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =10500
                            Top =780
                            Width =900
                            Height =240
                            Name ="Label53"
                            Caption ="&Print report"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6120
                    Top =720
                    Width =900
                    Height =720
                    TabIndex =12
                    Name ="cmdReassign"
                    Caption ="Re-Assign Item ID"
                    OnClick ="[Event Procedure]"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3960
                    Top =720
                    Width =900
                    Height =720
                    TabIndex =13
                    Name ="cmdPalmFrom"
                    Caption ="Command56"
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
                        0x0707070707070707070707070707070707000700000000000000000000000000 ,
                        0x0000070707070707070707070707070701f800ffffffffffffffffffffffffff ,
                        0xffff000707070707070707000000000701f800ffffffffffffffffffffffffff ,
                        0xffff000707070707070700f901f9010001f800f8f8f8f8f8f8f8f8f8fff800f8 ,
                        0xffff0007070707070700f901f901000001f8000000000000ffffffff0000ffff ,
                        0xffff00070707070700f901f9f900ffff0000ffffffffffff0000000001000707 ,
                        0x070700070707070700f9f9f900ffffffffffffffffffffff0707070000ffffff ,
                        0xffff00070707070700f9f9f900ffffffffffffffffff00000000000007070707 ,
                        0x070700070707070700f9f9f900ffffffffffffffff0700ff00fb0000ffffffff ,
                        0xfff80007070707070700f9f900ffffffffffffffff070000fb00070007070707 ,
                        0x0707000707070707070700000000ffffffffffffff0700fb0007ff00fffffff8 ,
                        0xf8f80007070707070707070707070000ffffffffff00fb0007ff000707070707 ,
                        0x0707000707070707070707070707070700ffffff00fb0007ff00fffffff8f8f8 ,
                        0xf8f8000707070707070707070707070701000000fb00ffff0007070707070707 ,
                        0x0707000707070707070707070707070701f800fb00000000fffffff8f8f8f8f8 ,
                        0xf8f800070707070707070707070707070100fb00070707070707070707070707 ,
                        0x0707000707070707070707070707070700fb00fffffffffffff8f8f8f8f8f8f8 ,
                        0xf8f80007070707070707070707070700fb000007070707070707070707070707 ,
                        0x070700070707070707070707070700fb00f800ffffff00000000000000000000 ,
                        0x0000000000000000070707070700fb0001f80007070700f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f900070707070700000701f800ffff0005f905f905f905f905f9 ,
                        0x05f905f905f90007070707070707070701f800070700f905f905f905f905f905 ,
                        0xf905f905f9050007070707070707070701f800ff00f905f905f905f905f905f9 ,
                        0x05f905f905000707070707070707070701f800070005f905f905f905f905f905 ,
                        0xf905f905f9000707070707070707070701f8000005f905f905f905f905f905f9 ,
                        0x05f905f900070707070707070707070701f80000f905f905f905f905f905f905 ,
                        0xf905f905000707070707070707070707010700f905f905f905f905f905f905f9 ,
                        0x05f9050007070707070707070707070707010005000001000001000001000001 ,
                        0x0000010007070707070707070707070707010100f80000f80000f80000f80000 ,
                        0xf800000707070707070707070707070707070707000007000007000007000007 ,
                        0x0000070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Run Query"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4920
                            Top =720
                            Width =1020
                            Height =660
                            Name ="Label57"
                            Caption ="Receive from Palm (Tab)"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2160
                    Top =720
                    Width =900
                    Height =720
                    TabIndex =14
                    Name ="cmdPalm"
                    Caption ="Command54"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada0000000000000dad0fffffffffff0ada ,
                        0x0fffffffffff0dad0fffffff770f0ada0fffffffff000dad0f888888880b0ada ,
                        0x0f777777770b0dad0f888888880bb0da0f777777770bb0ad0ffffffffff0bb0a ,
                        0x0f8888fffff0bb0d0f7777ffffff00000fffffff777f01100fffffffffff0000 ,
                        0x0000000000000dad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Run Query"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3180
                            Top =720
                            Width =660
                            Height =480
                            Name ="Label59"
                            Caption ="Send To Palm"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =65
                    Left =120
                    Top =720
                    Width =576
                    Height =576
                    TabIndex =15
                    Name ="cmdAddNewRecord"
                    Caption ="Command81"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddcdadadadadadadaaccdadadada4adaddcccdadad4d4d4da ,
                        0xaccccdadad444daddcccccda4444444aaccccdadad444daddcccdadad4d4d4da ,
                        0xaccdadadada4adaddcdadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Add Record"
                    UnicodeAccessKey =65

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =720
                            Top =720
                            Width =885
                            Height =435
                            Name ="Label82"
                            Caption ="&Add NEW Records"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "Audit.cls"
