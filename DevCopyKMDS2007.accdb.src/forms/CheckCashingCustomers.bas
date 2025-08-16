Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10080
    DatasheetFontHeight =10
    ItemSuffix =56
    Left =3795
    Top =945
    Right =13995
    Bottom =7320
    FrozenColumns =5
    HelpContextId =5500
    OnUnload ="[Event Procedure]"
    Tag ="22 Check Cashing Customers"
    RecSrcDt = Begin
        0x3890a5f80ae5e240
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="SELECT CheckCashing.* FROM CheckCashing ORDER BY CheckCashing.LastName; "
    Caption =" KMDS Software ©  KMDS Software ©  KMDS Software © "
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnError ="[Event Procedure]"
    OnMouseWheel ="[Event Procedure]"
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
            CanGrow = NotDefault
            Height =5340
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =8865
                    Top =180
                    Width =1080
                    Height =300
                    ColumnWidth =1440
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =8865
                    LayoutCachedTop =180
                    LayoutCachedWidth =9945
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =6840
                            Top =180
                            Width =1905
                            Height =285
                            Name ="ID_Label"
                            Caption ="Setting up a New Record"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =6840
                            LayoutCachedTop =180
                            LayoutCachedWidth =8745
                            LayoutCachedHeight =465
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =120
                    Width =9960
                    Height =5220
                    Name ="boxScreen"
                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =5340
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =120
                    Top =4800
                    Width =9960
                    Height =540
                    Name ="Box13"
                    LayoutCachedLeft =120
                    LayoutCachedTop =4800
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =5340
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    ColumnCount =5
                    ListRows =30
                    ListWidth =7200
                    Left =2445
                    Top =300
                    Width =2775
                    Height =300
                    TabIndex =1
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="cboID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CheckCashing.ID, CheckCashing.LastName, CheckCashing.FirstName, CheckCash"
                        "ing.Address, CheckCashing.City FROM CheckCashing ORDER BY CheckCashing.LastName;"
                        " "
                    ColumnWidths ="0;1440;1440;2160;1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =600
                            Top =300
                            Width =1710
                            Height =495
                            Name ="Label15"
                            Caption ="Select an Customer or Click on Add New "
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    Left =1380
                    Top =900
                    Width =2910
                    Height =300
                    TabIndex =2
                    Name ="FirstName"
                    ControlSource ="FirstName"
                    FontName ="Trebuchet MS"
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
                            Left =435
                            Top =900
                            Width =885
                            Height =285
                            Name ="Desc_Label"
                            Caption ="First Name"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    Left =1410
                    Top =1620
                    Width =2910
                    Height =300
                    TabIndex =4
                    Name ="Address1"
                    ControlSource ="Address"
                    FontName ="Trebuchet MS"
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
                            Left =540
                            Top =1620
                            Width =810
                            Height =285
                            Name ="Label23"
                            Caption ="Address1:"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    Left =1410
                    Top =1990
                    Width =1125
                    Height =315
                    TabIndex =5
                    Name ="City"
                    ControlSource ="City"
                    FontName ="Trebuchet MS"
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
                            Left =900
                            Top =1990
                            Width =450
                            Height =285
                            Name ="Label25"
                            Caption ="City:"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    Left =2930
                    Top =1990
                    Width =285
                    Height =285
                    TabIndex =6
                    Name ="St"
                    ControlSource ="State"
                    FontName ="Trebuchet MS"
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
                            Left =2605
                            Top =1990
                            Width =285
                            Height =285
                            Name ="Label26"
                            Caption ="St:"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    Left =3645
                    Top =1990
                    Width =645
                    Height =285
                    TabIndex =7
                    Name ="Zip"
                    ControlSource ="Zip"
                    FontName ="Trebuchet MS"
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
                            Left =3240
                            Top =1990
                            Width =360
                            Height =285
                            Name ="Label27"
                            Caption ="Zip:"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    Left =1410
                    Top =2410
                    Width =2910
                    Height =300
                    TabIndex =8
                    Name ="Phone1"
                    ControlSource ="TelephoneNumbr"
                    FontName ="Trebuchet MS"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
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
                            Left =675
                            Top =2410
                            Width =675
                            Height =285
                            Name ="Label28"
                            Caption ="Phone1:"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    Left =1395
                    Top =1260
                    Width =2910
                    Height =300
                    TabIndex =3
                    Name ="LastName"
                    ControlSource ="LastName"
                    FontName ="Trebuchet MS"
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
                            Left =420
                            Top =1305
                            Width =855
                            Height =285
                            Name ="Label46"
                            Caption ="Last Name"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =7140
                    Top =3240
                    Width =576
                    Height =576
                    TabIndex =9
                    Name ="Command45"
                    Caption ="Command45"
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
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707000000000000000000000000000000000000 ,
                        0x0000070707070707070707070700fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f80007070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8000000000000000000000000000000000000 ,
                        0x0000f80007070707070707070000ffffffffffffffffffffffffffffffffffff ,
                        0x07070000070707070707070700ffffffffffffff0000ff0000ff0000ffffffff ,
                        0xff07070007070707070707070700fffffffffffff800fff800fff800ffffffff ,
                        0xff07000707070707070707070700ffffffffffffffffffffffffffffffffffff ,
                        0xff07000707070707070707070700ffffffffffff0000ff0000ff0000ffffffff ,
                        0xff07000707070707070707070700fffffffffffff800fff800fff800ffffffff ,
                        0x070700070707070707070707070700ffffffffffffffffffffffffffffffffff ,
                        0x070007070707070707000000000000ffffffffff0000ff0000ff0000ffffffff ,
                        0x07000000000000070007070707f800fffffffffff800fff800fff800ffffff07 ,
                        0x0700070707f8f8000007070707f8f800ffffffffffffffffffffffffffffff07 ,
                        0x0007070707f8f8000000000000000000ffffffff0000ff0000ff0000ffffff07 ,
                        0x000000000000000000ffff0707070700fffffffff800fff800fff800ffff0707 ,
                        0x000707070707f80000ffffff0707070700ffffffffffffffffffffffffff0700 ,
                        0xf80707070707f8000700ffffff070707070000000700000000000007000000f8 ,
                        0x0707070707f800070700ffffffff07070707f8000700f8f8f8f8000700f8f807 ,
                        0x0707070707f8000707070000ffffffffff070707000707070707070007070707 ,
                        0x070707f800000707070707070000ffffffffff07070707070707070707070707 ,
                        0x070700000707070707070707070700000000ffffffffffffffffffffffff0000 ,
                        0x0000070707070707070707070707070707070000000000000000000000000707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x07070707070707070063010203022c000103011000010301ae03010301300701
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Auto Dialer"

                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1410
                    Top =3180
                    Height =300
                    TabIndex =10
                    Name ="Dob"
                    ControlSource ="Dob"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =930
                            Top =3180
                            Width =420
                            Height =240
                            Name ="Label47"
                            Caption ="Dob:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1395
                    Top =3540
                    Height =300
                    TabIndex =11
                    Name ="DriverLicense"
                    ControlSource ="DriverLicense"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =240
                            Top =3540
                            Width =1095
                            Height =240
                            Name ="Label48"
                            Caption ="DriverLicense:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1395
                    Top =3900
                    Width =405
                    Height =300
                    TabIndex =12
                    Name ="CheckCode"
                    ControlSource ="CheckCode"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =375
                            Top =3900
                            Width =960
                            Height =240
                            Name ="Label49"
                            Caption ="CheckCode:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    Left =2100
                    Top =3900
                    Width =4410
                    Height =240
                    Name ="Label50"
                    Caption ="\"X\" - Cash Register will not allow this customer to be selected"
                End
                Begin TextBox
                    OverlapFlags =247
                    Left =1440
                    Top =4320
                    Width =2910
                    Height =300
                    TabIndex =13
                    Name ="Text53"
                    ControlSource ="BankAccount"
                    FontName ="Trebuchet MS"
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
                            Left =360
                            Top =4320
                            Width =1020
                            Height =285
                            Name ="Label54"
                            Caption ="Reason for X"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    Left =4425
                    Top =4260
                    Width =4740
                    Height =420
                    Name ="Label55"
                    Caption ="This will pop up at the register when \"X\" code is encountered."
                End
            End
        End
        Begin FormFooter
            Height =1020
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =69
                    Left =3960
                    Top =60
                    Width =666
                    Height =606
                    Name ="cmdfrmDeleteRecord"
                    Caption ="Delete"
                    OnClick ="macKMDSDelete"
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
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"
                    UnicodeAccessKey =101
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3960
                    LayoutCachedTop =60
                    LayoutCachedWidth =4626
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =3998
                            Top =660
                            Width =555
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblDeleteRecord"
                            Caption ="D&elete"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =3998
                            LayoutCachedTop =660
                            LayoutCachedWidth =4553
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =86
                    Left =2880
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =1
                    Name ="cmdfrmSaveRecord"
                    Caption ="Command34"
                    OnClick ="macKMDSSave"
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
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Save Changes to Record"
                    UnicodeAccessKey =118
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =3546
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =2947
                            Top =660
                            Width =420
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblSaveRecord"
                            Caption ="Sa&ve"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =2947
                            LayoutCachedTop =660
                            LayoutCachedWidth =3367
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =78
                    Left =1800
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =2
                    Name ="cmdfrmUndo"
                    Caption ="Command35"
                    OnClick ="macKMDSUndo"
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
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1800
                    LayoutCachedTop =60
                    LayoutCachedWidth =2466
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =1867
                            Top =660
                            Width =450
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblUndo"
                            Caption ="U&ndo"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =1867
                            LayoutCachedTop =660
                            LayoutCachedWidth =2317
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =77
                    Left =7200
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =3
                    Name ="Command0"
                    Caption ="Command64"
                    OnClick ="macKMDSClose"
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
                    UnicodeAccessKey =109
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7200
                    LayoutCachedTop =60
                    LayoutCachedWidth =7866
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =7200
                            Top =660
                            Width =885
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label1"
                            Caption ="Close For&m"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =7200
                            LayoutCachedTop =660
                            LayoutCachedWidth =8085
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =65
                    Left =720
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =4
                    Name ="cmdfrmAddNew"
                    Caption ="Command67"
                    OnClick ="macKMDSAdd"
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
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Add a New Record"
                    UnicodeAccessKey =65
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =720
                    LayoutCachedTop =60
                    LayoutCachedWidth =1386
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =728
                            Top =660
                            Width =705
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label79"
                            Caption ="&Add New"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =728
                            LayoutCachedTop =660
                            LayoutCachedWidth =1433
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =5040
                    Top =60
                    Width =666
                    Height =606
                    FontSize =9
                    TabIndex =5
                    ForeColor =-2147483630
                    Name ="cmdPrintPreview"
                    Caption ="Show All Agents"
                    OnClick ="macKMDSPrintPreview"
                    FontName ="Trebuchet MS"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000030507050203040ff ,
                        0x203850500000000000000000b0a090ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff504840ff305060ff4078c0ff ,
                        0x304860ff3050605000000000b0a090ffffffffffb0a090ffb0a090ffb0a090ff ,
                        0xb0a090ffb0a090ffb0a090ffb0a090ff908880ff305070ff3080d0ff4098e0ff ,
                        0x50b0f0ff506870ff00000000b0a090fffffffffffffffffffff8fffff0f0f0ff ,
                        0xd0d8d0ff909890ff706860ff606860ff505860ff407090ff40a0e0ff60c8ffff ,
                        0x7090a0ff7088905000000000b0a090ffffffffffffffffffffffffffe0e0e0ff ,
                        0x909090ffb0a8a0ffd0c0b0ffd0b0a0ff807870ff505850ff6090b0ff7098b0ff ,
                        0x708890800000000000000000b0a090ffffffffffffffffffffffffffb0b0b0ff ,
                        0xc0b8b0fffff0e0ffffe8e0fff0d8c0fff0d0b0ff807870ff6058609070889080 ,
                        0x000000000000000000000000c0a890ffffffffffffffffffffffffffa09890ff ,
                        0xf0e8e0fffff8f0fffff0f0ffffe8e0fff0d8d0ffd0b0a0ff505850e000000000 ,
                        0x000000000000000000000000c0a8a0ffffffffffffffffffffffffffa0a090ff ,
                        0xf0e8e0fffffffffffff8f0fffff0f0ffffe8e0ffe0c0b0ff606060e000000000 ,
                        0x000000000000000000000000c0b0a0ffffffffffffffffffffffffffc0c8c0ff ,
                        0xc0c0c0fffffffffffffffffffff8f0fffff0e0ffb0a090ff6060608000000000 ,
                        0x000000000000000000000000d0b0a0fffffffffffffffffffffffffff0f8ffff ,
                        0xc0b8b0ffc0c0c0fff0e8e0fff0e8e0ffb0b0a0ff707060ff8070601000000000 ,
                        0x000000000000000000000000d0b8a0ffffffffffffffffffffffffffffffffff ,
                        0xf0f8ffffc0c8c0ffa0a090ff909080ff909090ff605040ff0000000000000000 ,
                        0x000000000000000000000000d0b8b0ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffb0a090ff604830ff604830ff604830ff0000000000000000 ,
                        0x000000000000000000000000d0c0b0ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffc0a890ffd0c8c0ff604830ffd0b0a0900000000000000000 ,
                        0x000000000000000000000000e0c0b0ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffc0a8a0ff604830ffd0b0a090000000000000000000000000 ,
                        0x000000000000000000000000e0c0b0ffe0c0b0ffe0c0b0ffe0c0b0ffe0c0b0ff ,
                        0xd0c0b0ffd0b8b0ffd0b0a0ffd0b0a09000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5040
                    LayoutCachedTop =60
                    LayoutCachedWidth =5706
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =5040
                            Top =660
                            Width =1095
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label61"
                            Caption ="Print Preview"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =5040
                            LayoutCachedTop =660
                            LayoutCachedWidth =6135
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =6120
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =6
                    Name ="cmdDataSheet"
                    Caption ="Command59"
                    OnClick ="macKMDSDatasheet"
                    ControlTipText ="Open Form"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000806860ff604830ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff604830ff604830ff604830ff ,
                        0x0000000000000000907860ffd0c8c0ff805840fff0f0f0ffe0d8d0ffd0d0d0ff ,
                        0xd0c8c0ffd0c0c0ffc0b8b0ffc0b0b0ffb0a8a0ffb0a8a0ffa09890ff604830ff ,
                        0x0000000000000000a08070ffd0c8d0ff805840fff0f0f0ffb09080ffa08070ff ,
                        0xe0e0d0ff907050ff906850ffd0c8c0ff906850ff906850ffb09890ff604830ff ,
                        0x0000000000000000a09080ffd0d0d0ff805840fffff8f0fff0f0f0fff0f0f0ff ,
                        0xf0e8e0ffe0e0e0ffe0d8d0ffe0d0d0ffd0c8c0ffd0c0b0ffb0a8a0ff604830ff ,
                        0x0000000000000000b0a090ffd0d0d0ff906040fffff8ffffc0a8a0ffc09890ff ,
                        0xf0f0f0ffa08070ffa07860ffe0d8d0ff906850ff906850ffc0b0b0ff604830ff ,
                        0x0000000000000000c0a8a0ffe0d8e0ff906850fffffffffffff8fffffff8f0ff ,
                        0xf0f0f0fff0f0f0fff0e8e0ffe0e0e0ffe0d8d0ffe0d8d0ffc0b8b0ff604830ff ,
                        0x0000000000000000c0b0a0ffe0e0e0ffa07050ffffffffffc0a8a0ffc0a090ff ,
                        0xfff8f0ffc09890ffb09080fff0e8e0ffa08060ffa07860ffd0c8c0ff604830ff ,
                        0x0000000000000000c0b8a0ffe0e0e0ffa07860ffffffffffffffffffffffffff ,
                        0xfff8fffffff8f0fff0f0f0fff0f0f0fff0e8e0ffe0e0e0ffd0d0d0ff604830ff ,
                        0x0000000000000000d0b8b0fff0e8f0ffb08060ffffffffffc0a890ffc0a090ff ,
                        0xffffffffc0a090ffc0a090fff0f0f0ffb09080ffb09080ffe0d8d0ff604830ff ,
                        0x0000000000000000d0c0b0fff0f0f0ffb08870ffffffffffffffffffffffffff ,
                        0xfffffffffffffffffff8fffffff8f0fff0f0f0fff0f0f0ffe0e0e0ff604830ff ,
                        0x0000000000000000d0c0b0ffc09080ffb09070ffb08870ffa08060ffa07860ff ,
                        0x907050ff906850ff906040ff806040ff805840ff805840ff805840ff604830ff ,
                        0x0000000000000000d0c0b0fff0f0f0ffc09080fff0f0f0fff0e8f0ffe0e8e0ff ,
                        0xe0e0e0ffe0e0e0ffe0d8e0ffd0d8d0ffd0d0d0ffd0d0d0ffd0d0d0ff604830ff ,
                        0x0000000000000000d0c0b0ffd0c0b0ffd0c0b0ffd0c0b0ffd0b8b0ffc0b0a0ff ,
                        0xc0b0a0ffc0a8a0ffb0a090ffb09890ffa09080ffa08870ff908070ff907860ff ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6120
                    LayoutCachedTop =60
                    LayoutCachedWidth =6786
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =2
                            Left =6120
                            Top =660
                            Width =915
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label60"
                            Caption ="Datasheet"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =6120
                            LayoutCachedTop =660
                            LayoutCachedWidth =7035
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =215
                    Left =120
                    Width =9960
                    Height =960
                    Name ="Box81"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =960
                End
            End
        End
    End
End
CodeBehindForm
' See "CheckCashingCustomers.cls"
