Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =255
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10710
    DatasheetFontHeight =10
    ItemSuffix =32
    Top =285
    Right =10110
    Bottom =9600
    HelpContextId =5400
    AfterInsert ="[Event Procedure]"
    AfterDelConfirm ="[Event Procedure]"
    Tag ="64 Kit Sales"
    Filter ="((ID>1))"
    RecSrcDt = Begin
        0xe535eb73aae7e340
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="Kit"
    Caption ="Kit KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Sof"
        "tware © "
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    OnMouseWheel ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextAlign =3
            FontSize =9
            FontName ="Trebuchet MS"
        End
        Begin Rectangle
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =9
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Trebuchet MS"
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
            FontSize =9
            FontName ="Trebuchet MS"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontSize =9
            FontName ="Trebuchet MS"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontSize =9
            FontName ="Trebuchet MS"
        End
        Begin Subform
            OldBorderStyle =0
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin ToggleButton
            FontSize =9
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Trebuchet MS"
            BorderLineStyle =0
        End
        Begin Tab
            FontSize =9
            FontName ="Trebuchet MS"
            BorderLineStyle =0
        End
        Begin FormHeader
            Visible = NotDefault
            Height =471
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5199
                    Height =330
                    Name ="ItemID"
                    ControlSource ="ItemID"

                    LayoutCachedLeft =5199
                    LayoutCachedWidth =6639
                    LayoutCachedHeight =330
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4080
                            Top =15
                            Width =1050
                            Height =315
                            Name ="Label16"
                            Caption =" Kit Item"
                            LayoutCachedLeft =4080
                            LayoutCachedTop =15
                            LayoutCachedWidth =5130
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1854
                    Height =330
                    TabIndex =1
                    Name ="ID"
                    ControlSource ="ID"

                    LayoutCachedLeft =1854
                    LayoutCachedWidth =3294
                    LayoutCachedHeight =330
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1500
                            Width =285
                            Height =315
                            Name ="Label17"
                            Caption ="ID"
                            LayoutCachedLeft =1500
                            LayoutCachedWidth =1785
                            LayoutCachedHeight =315
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7545
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1869
                    Top =900
                    Width =1701
                    Height =330
                    ColumnWidth =1935
                    TabIndex =3
                    Name ="Price"
                    ControlSource ="Price"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =1869
                    LayoutCachedTop =900
                    LayoutCachedWidth =3570
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =540
                            Top =900
                            Width =1260
                            Height =315
                            Name ="Description_Label"
                            Caption ="Kit Price"
                            LayoutCachedLeft =540
                            LayoutCachedTop =900
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =1215
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    ColumnHidden = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =7
                    ListWidth =8640
                    Left =1914
                    Top =472
                    Width =2070
                    Height =330
                    ColumnWidth =1440
                    BoundColumn =3
                    Name ="cboID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryKitsOnFile.Description, qryKitsOnFile.Price, qryKitsOnFile.ItemID, qry"
                        "KitsOnFile.ID, qryKitsOnFile.RetailPrice, qryKitsOnFile.Discount, IIf([RetailPri"
                        "ce]=0,0,Format([discount]/[RetailPrice],\"Percent\")) AS DiscPercent FROM qryKit"
                        "sOnFile; "
                    ColumnWidths ="2835;1134;1134;0"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =1914
                    LayoutCachedTop =472
                    LayoutCachedWidth =3984
                    LayoutCachedHeight =802
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =720
                            Top =472
                            Width =1125
                            Height =320
                            Name ="Price Group_Label"
                            Caption ="Kit"
                            EventProcPrefix ="Price_Group_Label"
                            LayoutCachedLeft =720
                            LayoutCachedTop =472
                            LayoutCachedWidth =1845
                            LayoutCachedHeight =792
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =120
                    Top =6900
                    Width =10575
                    Height =645
                    Name ="Box13"
                    LayoutCachedLeft =120
                    LayoutCachedTop =6900
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =7545
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =735
                    Top =7080
                    Width =1830
                    Height =285
                    ColumnWidth =1440
                    FontSize =8
                    TabIndex =9
                    BackColor =12632256
                    ForeColor =-2147483640
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =735
                    LayoutCachedTop =7080
                    LayoutCachedWidth =2565
                    LayoutCachedHeight =7365
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =180
                            Top =7080
                            Width =495
                            Height =240
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="SetupDate_Label"
                            Caption ="Setup"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =180
                            LayoutCachedTop =7080
                            LayoutCachedWidth =675
                            LayoutCachedHeight =7320
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =8730
                    Top =7080
                    Width =1830
                    Height =285
                    ColumnWidth =1440
                    FontSize =8
                    TabIndex =8
                    BackColor =12632256
                    ForeColor =-2147483640
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    Format ="General Date"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =8730
                    LayoutCachedTop =7080
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =7365
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =8040
                            Top =7080
                            Width =630
                            Height =270
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="ChangeDate_Label"
                            Caption ="Change"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =8040
                            LayoutCachedTop =7080
                            LayoutCachedWidth =8670
                            LayoutCachedHeight =7350
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OldBorderStyle =1
                    OverlapFlags =255
                    Left =60
                    Width =10590
                    Height =7545
                    Name ="boxScreen"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =10650
                    LayoutCachedHeight =7545
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =1
                    Left =4935
                    Top =6075
                    Width =5055
                    Height =585
                    Name ="Label13"
                    Caption ="The Discount is the difference of the Kit Price and Retail Value.  The discount "
                        "will be rung up to the above Kit item. "
                    LayoutCachedLeft =4935
                    LayoutCachedTop =6075
                    LayoutCachedWidth =9990
                    LayoutCachedHeight =6660
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    AccessKey =68
                    ColumnCount =7
                    ListRows =25
                    ListWidth =7560
                    Left =7679
                    Top =420
                    Width =1650
                    Height =330
                    ColumnWidth =3645
                    TabIndex =2
                    BoundColumn =2
                    Name ="cmbItemByDesc"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="lkupFormItemByDescription"
                    ColumnWidths ="2881;720;720;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    UnicodeAccessKey =68

                    LayoutCachedLeft =7679
                    LayoutCachedTop =420
                    LayoutCachedWidth =9329
                    LayoutCachedHeight =750
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =162
                            Left =6545
                            Top =420
                            Width =1065
                            Height =285
                            Name ="Label14"
                            Caption ="&Description"
                            LayoutCachedLeft =6545
                            LayoutCachedTop =420
                            LayoutCachedWidth =7610
                            LayoutCachedHeight =705
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    ColumnCount =8
                    ListRows =25
                    ListWidth =9000
                    Left =4854
                    Top =439
                    Width =1410
                    Height =330
                    TabIndex =1
                    BoundColumn =3
                    Name ="cmbItemByUPC"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="lkupFormItemByUPC"
                    ColumnWidths ="1441;2881;720;720;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =4854
                    LayoutCachedTop =439
                    LayoutCachedWidth =6264
                    LayoutCachedHeight =769
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =162
                            Left =4110
                            Top =439
                            Width =675
                            Height =285
                            Name ="Label15"
                            Caption ="Kit Item"
                            LayoutCachedLeft =4110
                            LayoutCachedTop =439
                            LayoutCachedWidth =4785
                            LayoutCachedHeight =724
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =247
                    Left =300
                    Top =1624
                    Width =10140
                    Height =3360
                    TabIndex =7
                    Name ="KitISubForm"
                    SourceObject ="Form.KitISubForm"
                    LinkChildFields ="KitID"
                    LinkMasterFields ="ID"

                    LayoutCachedLeft =300
                    LayoutCachedTop =1624
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =4984
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =0
                            Left =300
                            Top =1320
                            Width =3210
                            Height =345
                            Name ="KitISubForm Label"
                            Caption ="Items sold when the Kit is purchased."
                            EventProcPrefix ="KitISubForm_Label"
                            LayoutCachedLeft =300
                            LayoutCachedTop =1320
                            LayoutCachedWidth =3510
                            LayoutCachedHeight =1665
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1824
                    Top =5730
                    Width =1701
                    Height =315
                    FontSize =11
                    TabIndex =4
                    BackColor =255
                    BorderColor =5137573
                    Name ="txtSumPrice"
                    ControlSource ="=IIf(IsNull([KitISubForm].[Form]![txtSumPrice]),0,[KitISubForm].[Form]![txtSumPr"
                        "ice])"
                    Format ="Fixed"
                    FontName ="Calibri"

                    LayoutCachedLeft =1824
                    LayoutCachedTop =5730
                    LayoutCachedWidth =3525
                    LayoutCachedHeight =6045
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =765
                            Top =5730
                            Width =1065
                            Height =300
                            Name ="Label18"
                            Caption ="Retail Value"
                            LayoutCachedLeft =765
                            LayoutCachedTop =5730
                            LayoutCachedWidth =1830
                            LayoutCachedHeight =6030
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1824
                    Top =6135
                    Width =1701
                    Height =315
                    FontSize =11
                    TabIndex =5
                    BackColor =255
                    BorderColor =5137573
                    Name ="txtDiscount"
                    ControlSource ="=[txtSumPrice]-[Price]"
                    Format ="Fixed"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =1824
                    LayoutCachedTop =6135
                    LayoutCachedWidth =3525
                    LayoutCachedHeight =6450
                    ConditionalFormat14 = Begin
                        0x01000100000000000000050000000100000000000000ed1c2400010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =720
                            Top =6135
                            Width =1035
                            Height =300
                            Name ="Label20"
                            Caption ="Discount:"
                            LayoutCachedLeft =720
                            LayoutCachedTop =6135
                            LayoutCachedWidth =1755
                            LayoutCachedHeight =6435
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =1
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3759
                    Top =6150
                    Width =1011
                    Height =315
                    FontSize =11
                    TabIndex =6
                    BackColor =255
                    BorderColor =5137573
                    Name ="txtDiscPercent"
                    ControlSource ="=([txtSumPrice]-[Price])/[Price]"
                    Format ="Percent"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =3759
                    LayoutCachedTop =6150
                    LayoutCachedWidth =4770
                    LayoutCachedHeight =6465
                    ConditionalFormat14 = Begin
                        0x01000100000000000000050000000100000000000000ed1c2400010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =3660
                    Top =904
                    Width =6495
                    Height =315
                    Name ="Label23"
                    Caption ="The Kit Price is the amount the customer will pay for the items listed below."
                    LayoutCachedLeft =3660
                    LayoutCachedTop =904
                    LayoutCachedWidth =10155
                    LayoutCachedHeight =1219
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =1
                    Left =3840
                    Top =5220
                    Width =6135
                    Height =660
                    Name ="Label24"
                    Caption ="The Retail value is the total of the Retail price of the items listed above.  Bu"
                        "ydown discounts are NOT displayed."
                    LayoutCachedLeft =3840
                    LayoutCachedTop =5220
                    LayoutCachedWidth =9975
                    LayoutCachedHeight =5880
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =720
                    Top =60
                    Width =6495
                    Height =315
                    Name ="lblInitialInstruction"
                    Caption ="Select a Kit to change   OR    Click the 'Add New' To set up a new Kit."
                    LayoutCachedLeft =720
                    LayoutCachedTop =60
                    LayoutCachedWidth =7215
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1809
                    Top =5280
                    Width =1701
                    Height =315
                    FontSize =11
                    TabIndex =10
                    BackColor =255
                    BorderColor =5137573
                    Name ="txtLiquorValue"
                    ControlSource ="=IIf(IsNull([KitISubForm].[Form]![txtSumLiquor]),0,[KitISubForm].[Form]![txtSumL"
                        "iquor])"
                    Format ="Fixed"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =1809
                    LayoutCachedTop =5280
                    LayoutCachedWidth =3510
                    LayoutCachedHeight =5595
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =750
                            Top =5280
                            Width =1065
                            Height =300
                            Name ="Label28"
                            Caption ="Liquor Value"
                            LayoutCachedLeft =750
                            LayoutCachedTop =5280
                            LayoutCachedWidth =1815
                            LayoutCachedHeight =5580
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OldBorderStyle =1
                    OverlapFlags =247
                    Left =120
                    Top =5100
                    Width =10560
                    Height =1740
                    Name ="Box31"
                    LayoutCachedLeft =120
                    LayoutCachedTop =5100
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =6840
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
                    Left =3480
                    Top =60
                    Width =666
                    Height =606
                    FontSize =8
                    ForeColor =0
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
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"
                    UnicodeAccessKey =101
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3480
                    LayoutCachedTop =60
                    LayoutCachedWidth =4146
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =3518
                            Top =660
                            Width =555
                            Height =285
                            FontSize =8
                            Name ="lblDeleteRecord"
                            Caption ="D&elete"
                            LayoutCachedLeft =3518
                            LayoutCachedTop =660
                            LayoutCachedWidth =4073
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =86
                    Left =2400
                    Top =60
                    Width =666
                    Height =606
                    FontSize =8
                    TabIndex =1
                    ForeColor =0
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
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Save Changes to Record"
                    UnicodeAccessKey =118
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2400
                    LayoutCachedTop =60
                    LayoutCachedWidth =3066
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =2467
                            Top =660
                            Width =420
                            Height =285
                            FontSize =8
                            Name ="lblSaveRecord"
                            Caption ="Sa&ve"
                            LayoutCachedLeft =2467
                            LayoutCachedTop =660
                            LayoutCachedWidth =2887
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =78
                    Left =1320
                    Top =60
                    Width =666
                    Height =606
                    FontSize =8
                    TabIndex =2
                    ForeColor =0
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
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Undo Changes Made to This Record"
                    UnicodeAccessKey =110
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1320
                    LayoutCachedTop =60
                    LayoutCachedWidth =1986
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =1387
                            Top =660
                            Width =450
                            Height =285
                            FontSize =8
                            Name ="lblUndo"
                            Caption ="U&ndo"
                            LayoutCachedLeft =1387
                            LayoutCachedTop =660
                            LayoutCachedWidth =1837
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =77
                    Left =6720
                    Top =60
                    Width =666
                    Height =606
                    FontSize =8
                    TabIndex =3
                    ForeColor =0
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
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"
                    UnicodeAccessKey =109
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6720
                    LayoutCachedTop =60
                    LayoutCachedWidth =7386
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =6720
                            Top =660
                            Width =885
                            Height =285
                            FontSize =8
                            Name ="Label1"
                            Caption ="Close For&m"
                            LayoutCachedLeft =6720
                            LayoutCachedTop =660
                            LayoutCachedWidth =7605
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =65
                    Left =240
                    Top =60
                    Width =666
                    Height =606
                    FontSize =8
                    TabIndex =4
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
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =906
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =248
                            Top =660
                            Width =705
                            Height =285
                            FontSize =8
                            Name ="Label79"
                            Caption ="&Add New"
                            LayoutCachedLeft =248
                            LayoutCachedTop =660
                            LayoutCachedWidth =953
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =4560
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =5
                    Name ="cmdPrintPreview"
                    Caption ="Show All Agents"
                    OnClick ="macKMDSPrintPreview"
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

                    LayoutCachedLeft =4560
                    LayoutCachedTop =60
                    LayoutCachedWidth =5226
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =4560
                            Top =660
                            Width =1095
                            Height =285
                            FontSize =8
                            Name ="Label61"
                            Caption ="Print Preview"
                            LayoutCachedLeft =4560
                            LayoutCachedTop =660
                            LayoutCachedWidth =5655
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =5640
                    Top =60
                    Width =666
                    Height =606
                    FontSize =8
                    TabIndex =6
                    ForeColor =0
                    Name ="cmdDataSheet"
                    Caption ="Command59"
                    OnClick ="macKMDSDatasheet"
                    FontName ="MS Sans Serif"
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

                    LayoutCachedLeft =5640
                    LayoutCachedTop =60
                    LayoutCachedWidth =6306
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =2
                            Left =5640
                            Top =660
                            Width =915
                            Height =285
                            FontSize =8
                            Name ="Label60"
                            Caption ="Datasheet"
                            LayoutCachedLeft =5640
                            LayoutCachedTop =660
                            LayoutCachedWidth =6555
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    Left =120
                    Width =7740
                    Height =960
                    Name ="Box81"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =960
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =162
                    TextAlign =1
                    Left =7980
                    Top =45
                    Width =2595
                    Height =855
                    FontSize =14
                    ForeColor =12349952
                    Name ="lblStatus"
                    Caption ="Changing an existing Kit."
                    LayoutCachedLeft =7980
                    LayoutCachedTop =45
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =900
                End
            End
        End
    End
End
CodeBehindForm
' See "Kit.cls"
