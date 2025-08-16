Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7875
    DatasheetFontHeight =10
    ItemSuffix =34
    Left =5895
    Top =2490
    Right =14100
    Bottom =6795
    HelpContextId =5300
    AfterInsert ="[Event Procedure]"
    AfterDelConfirm ="[Event Procedure]"
    Tag ="20 Size List"
    OrderBy ="Oz"
    RecSrcDt = Begin
        0xc6acc9665095e240
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="SELECT Size.* FROM [Size] ORDER BY Size.Description; "
    Caption ="  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Softw"
        "are ©  KMDS Software ©  KMDS Software ©  KMDS Software © "
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnDblClick ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    OnMouseWheel ="[Event Procedure]"
    FilterOnLoad =0
    SplitFormOrientation =1
    SplitFormSplitterBar =0
    SplitFormSize =4665
    SplitFormOrientation =1
    SplitFormSplitterBar =0
    SplitFormSize =4665
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    FitToScreen =255
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
            Height =3300
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    Left =2100
                    Top =300
                    Width =360
                    Height =255
                    ColumnWidth =1440
                    Name ="ID"
                    ControlSource ="ID"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =300
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =555
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =1020
                            Top =120
                            Width =975
                            Height =420
                            Name ="ID_Label"
                            Caption ="Setting up a New Record"
                            LayoutCachedLeft =1020
                            LayoutCachedTop =120
                            LayoutCachedWidth =1995
                            LayoutCachedHeight =540
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    Left =2055
                    Top =1020
                    Width =1935
                    Height =255
                    ColumnWidth =1935
                    TabIndex =2
                    Name ="Desc"
                    ControlSource ="Description"
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
                            Left =780
                            Top =1020
                            Width =1215
                            Height =240
                            Name ="Desc_Label"
                            Caption ="Size Description"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    Left =2055
                    Top =1380
                    Width =900
                    Height =255
                    ColumnWidth =3705
                    TabIndex =3
                    Name ="Oz"
                    ControlSource ="Oz"
                    StatusBarText ="How many ounces in each item."
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
                            Left =1725
                            Top =1380
                            Width =270
                            Height =240
                            Name ="Oz_Label"
                            Caption ="Oz"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    Left =2055
                    Top =1740
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =4
                    Name ="CaseQty"
                    ControlSource ="CaseQty"
                    StatusBarText ="Case quantity"
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
                            Left =1320
                            Top =1740
                            Width =675
                            Height =240
                            Name ="CaseQty_Label"
                            Caption ="CaseQty"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =120
                    Width =7740
                    Height =3180
                    Name ="boxScreen"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =2760
                    Width =7740
                    Height =540
                    Name ="Box13"
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    ColumnCount =2
                    ListRows =30
                    ListWidth =1440
                    Left =2085
                    Top =660
                    ColumnWidth =1440
                    TabIndex =1
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="cboID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Size.ID, Size.Description FROM [Size] ORDER BY Size.Description; "
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =2085
                    LayoutCachedTop =660
                    LayoutCachedWidth =3525
                    LayoutCachedHeight =900
                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =1020
                            Top =660
                            Width =1005
                            Height =240
                            Name ="Label15"
                            Caption ="Select a Size"
                            LayoutCachedLeft =1020
                            LayoutCachedTop =660
                            LayoutCachedWidth =2025
                            LayoutCachedHeight =900
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =735
                    Top =2940
                    Width =2535
                    Height =255
                    ColumnWidth =1440
                    TabIndex =5
                    BackColor =12632256
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                    LayoutCachedLeft =735
                    LayoutCachedTop =2940
                    LayoutCachedWidth =3270
                    LayoutCachedHeight =3195
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =180
                            Top =2940
                            Width =495
                            Height =240
                            Name ="SetupDate_Label"
                            Caption ="Setup"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =5670
                    Top =2940
                    Width =2040
                    Height =255
                    ColumnWidth =1440
                    TabIndex =6
                    BackColor =12632256
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    Format ="General Date"

                    LayoutCachedLeft =5670
                    LayoutCachedTop =2940
                    LayoutCachedWidth =7710
                    LayoutCachedHeight =3195
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =4980
                            Top =2940
                            Width =630
                            Height =240
                            Name ="ChangeDate_Label"
                            Caption ="Change"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =5520
                    Top =960
                    Width =576
                    Height =576
                    TabIndex =7
                    Name ="Command29"
                    Caption ="Command29"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Report to File"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000c08060f0905030ff905830ff904820ff ,
                        0x804020ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000c08060f0905030ffc07860ff ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000c08870ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000008000000080000000800000008000000080000000800000008000000080 ,
                        0x0000000000000000000000000000000000000000106010ff106010ff106010ff ,
                        0xc0ffe0ffc0ffe0ffc0ffe0ffc0ffe0ffc0ffe0ffc0ffe0ffc0ffe0ff00000080 ,
                        0x0000000000000000000000000000000000000000207020ff209020ff209020ff ,
                        0xd0ffe0ffd0ffe0ffd0ffe0ffc0e8d0ff609870ff407050ffd0ffe0ff00000080 ,
                        0x0000000000000000000000000000000000000000207820ff207810ff207810ff ,
                        0xe0fff0ff408850ff408050ff90c090ff60a860ff408050ffe0fff0ff00000080 ,
                        0x0000000000000000000000000000000000000000308030ff106810ff106810ff ,
                        0xe0fff0ffd0f0d0ff80b870ff60b060ff307850ffc0e0c0ffe0fff0ff00000080 ,
                        0x0000000000000000000000000000000000000000309030ff70b870ff70b870ff ,
                        0xf0fff0ffe0f8e0ff60b060ff409040ff508850ffe0f0e0fff0fff0ff00000080 ,
                        0x0000000000000000000000000000000000000000409840ff80c880ff80c880ff ,
                        0xffffffff60a860ff70a870ffffffffff90b880ff409040ffffffffff00000080 ,
                        0x000000000000000000000000000000000000000040a840ff90d090ffa0d0a0ff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff00000080 ,
                        0x000000000000000000000000000000000000000040b8508070c070ffa0d8a0ff ,
                        0xb0e0b0ffb0e0b0ffb0e0b0ffb0e0b0ffb0e0b0ff40b040ff0000000000000000 ,
                        0x00000000000000000000000000000000000000000000000040c0608050b850ff ,
                        0x50b850ff50b850ff50b850ff50b850ff50b850ff50b850ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5520
                    LayoutCachedTop =960
                    LayoutCachedWidth =6096
                    LayoutCachedHeight =1536
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =5040
                    Top =420
                    TabIndex =8
                    Name ="Text30"
                    ControlSource ="ID"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =5040
                    LayoutCachedTop =420
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4680
                            Top =420
                            Width =285
                            Height =240
                            Name ="Label31"
                            Caption ="ID:"
                            LayoutCachedLeft =4680
                            LayoutCachedTop =420
                            LayoutCachedWidth =4965
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin OptionButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =6600
                    Top =450
                    TabIndex =9
                    Name ="optNewLiquorSize"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="False"

                    LayoutCachedLeft =6600
                    LayoutCachedTop =450
                    LayoutCachedWidth =6860
                    LayoutCachedHeight =690
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =6825
                            Top =420
                            Width =870
                            Height =405
                            Name ="Label33"
                            Caption ="New Liquor Size"
                            LayoutCachedLeft =6825
                            LayoutCachedTop =420
                            LayoutCachedWidth =7695
                            LayoutCachedHeight =825
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
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =69
                    Left =3600
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

                    LayoutCachedLeft =3600
                    LayoutCachedTop =60
                    LayoutCachedWidth =4266
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =3638
                            Top =660
                            Width =555
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblDeleteRecord"
                            Caption ="D&elete"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =3638
                            LayoutCachedTop =660
                            LayoutCachedWidth =4193
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =86
                    Left =2520
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

                    LayoutCachedLeft =2520
                    LayoutCachedTop =60
                    LayoutCachedWidth =3186
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =2587
                            Top =660
                            Width =420
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblSaveRecord"
                            Caption ="Sa&ve"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =2587
                            LayoutCachedTop =660
                            LayoutCachedWidth =3007
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =78
                    Left =1440
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

                    LayoutCachedLeft =1440
                    LayoutCachedTop =60
                    LayoutCachedWidth =2106
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =1507
                            Top =660
                            Width =450
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblUndo"
                            Caption ="U&ndo"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =1507
                            LayoutCachedTop =660
                            LayoutCachedWidth =1957
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =77
                    Left =6840
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

                    LayoutCachedLeft =6840
                    LayoutCachedTop =60
                    LayoutCachedWidth =7506
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =6840
                            Top =660
                            Width =885
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label1"
                            Caption ="Close For&m"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =6840
                            LayoutCachedTop =660
                            LayoutCachedWidth =7725
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =65
                    Left =360
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

                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =1026
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =368
                            Top =660
                            Width =705
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label79"
                            Caption ="&Add New"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =368
                            LayoutCachedTop =660
                            LayoutCachedWidth =1073
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =4680
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

                    LayoutCachedLeft =4680
                    LayoutCachedTop =60
                    LayoutCachedWidth =5346
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =4680
                            Top =660
                            Width =1095
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label61"
                            Caption ="Print Preview"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =4680
                            LayoutCachedTop =660
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =5760
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

                    LayoutCachedLeft =5760
                    LayoutCachedTop =60
                    LayoutCachedWidth =6426
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =2
                            Left =5760
                            Top =660
                            Width =915
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label60"
                            Caption ="Datasheet"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =5760
                            LayoutCachedTop =660
                            LayoutCachedWidth =6675
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =215
                    Left =120
                    Width =7740
                    Height =960
                    Name ="Box81"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =960
                End
            End
        End
    End
End
CodeBehindForm
' See "Sizes.cls"
