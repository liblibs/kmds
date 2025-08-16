Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
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
    Width =13380
    DatasheetFontHeight =10
    ItemSuffix =41
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x63c1ea9d4fdee240
    End
    Caption =" KMDS Software ©  KMDS Software © "
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
            Height =1200
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =60
                    Width =4200
                    Height =480
                    FontSize =16
                    Name ="Label15"
                    Caption ="Inventory Adjustment Screen"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =5925
                    Top =420
                    Width =1875
                    BackColor =10092543
                    Name ="cmbListType"
                    RowSourceType ="Value List"
                    RowSource ="Code;Description;SalesDept;Vendor;PriceGroup;Location;Category"
                    ColumnWidths ="1440"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5040
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
                    Left =8760
                    Top =420
                    Width =1860
                    TabIndex =1
                    Name ="cmbSubCategoryID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.Description, Item.ID, Item.ID FROM Item ORDER BY Item.Description, I"
                        "tem.ID; "
                    ColumnWidths ="2880;144;144"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7860
                            Top =420
                            Width =840
                            Height =240
                            Name ="Description_Label"
                            Caption ="Code"
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
                    Left =8760
                    Top =720
                    Width =1860
                    TabIndex =2
                    Name ="cmbSubCategoryIDEnd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.Description, Item.ID, Item.ID FROM Item ORDER BY Item.Description, I"
                        "tem.ID; "
                    ColumnWidths ="2880;144;144"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8445
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
                    Left =10740
                    Top =420
                    Width =2586
                    Height =576
                    TabIndex =3
                    Name ="cmdAddItems"
                    Caption ="Add items from this range of Codes."
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Run Query"

                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =240
                    Top =720
                    Width =4142
                    Height =358
                    TabIndex =4
                    Name ="Frame24"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =360
                            Top =600
                            Width =705
                            Height =240
                            BackColor =-2147483633
                            Name ="Label25"
                            Caption ="Sort by"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =1260
                            Top =808
                            OptionValue =1
                            Name ="Option27"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1485
                                    Top =780
                                    Width =615
                                    Height =240
                                    Name ="Label28"
                                    Caption ="Item ID"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =2220
                            Top =808
                            OptionValue =2
                            Name ="Option29"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =2450
                                    Top =780
                                    Width =870
                                    Height =240
                                    Name ="Label30"
                                    Caption ="Description"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =3420
                            Top =808
                            OptionValue =3
                            Name ="Option31"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =3650
                                    Top =780
                                    Width =360
                                    Height =240
                                    Name ="Label32"
                                    Caption ="Size"
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =5100
                    Top =120
                    Width =8055
                    Height =210
                    Name ="Label33"
                    Caption ="Fill the list of items by selecting a range of items with the options below..."
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4980
                    Top =720
                    Width =1470
                    Height =405
                    TabIndex =5
                    Name ="cmdAuditUpdateOnShelf"
                    Caption ="Refresh OnShelf "
                    OnClick ="[Event Procedure]"

                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5505
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =215
                    Left =120
                    Top =120
                    Width =13260
                    Height =5385
                    Name ="AuditSub"
                    SourceObject ="Form.AuditSub"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =480
                            Width =720
                            Height =240
                            Name ="Label0"
                            Caption ="AuditSub"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1005
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
                    Left =3480
                    Top =360
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
                            Left =3480
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
                    Left =8160
                    Top =360
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
                            Left =8160
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
                    Left =5040
                    Top =360
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="Text7"
                    ControlSource ="=AuditSub.Form!TotVarianceAmount"
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
                            Left =5040
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
                    Left =6600
                    Top =360
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="Text9"
                    ControlSource ="=AuditSub.Form!TotExtPrice"
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
                            Left =6600
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
                    Left =11400
                    Top =180
                    Width =975
                    Height =825
                    TabIndex =5
                    Name ="cmdPost"
                    Caption ="Post This Batch"
                    OnClick ="[Event Procedure]"

                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =77
                    Left =12480
                    Top =180
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
                            OverlapFlags =87
                            TextAlign =2
                            Left =12480
                            Top =780
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
                    Left =60
                    Top =180
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

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =780
                            Width =1425
                            Height =210
                            Name ="Label22"
                            Caption ="Delete ALL records"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10620
                    Top =180
                    Width =576
                    Height =576
                    TabIndex =8
                    Name ="cmdPrint"
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

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =10140
                            Top =780
                            Width =1080
                            Height =210
                            Name ="Label37"
                            Caption ="Print report"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10020
                    Top =540
                    Width =245
                    TabIndex =9
                    Name ="chkOnlyVariance"

                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextAlign =2
                            Left =9660
                            Top =60
                            Width =900
                            Height =480
                            Name ="Label40"
                            Caption ="Only Print Variance"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "Auditx.cls"
