Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9060
    DatasheetFontHeight =10
    ItemSuffix =19
    Left =4305
    Right =13545
    Bottom =9675
    HelpContextId =5200
    Tag ="61 Category List"
    OrderBy ="Category.ID"
    RecSrcDt = Begin
        0x708297189696e240
    End
    RecordSource ="Category"
    Caption ="Category"
    OnCurrent ="[Event Procedure]"
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =8700
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1920
                    Top =540
                    Width =945
                    Height =330
                    ColumnWidth =945
                    TabIndex =1
                    Name ="ID"
                    ControlSource ="ID"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1620
                            Top =540
                            Width =225
                            Height =315
                            Name ="ID_Label"
                            Caption ="ID"
                        End
                    End
                End
                Begin TextBox
                    ScrollBars =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1920
                    Top =900
                    Width =4560
                    Height =330
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="Description"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =840
                            Top =900
                            Width =1020
                            Height =315
                            Name ="Description_Label"
                            Caption ="Description"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1620
                    Top =120
                    Width =4620
                    Height =330
                    BackColor =16777164
                    Name ="cmbDescription"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Category.ID, Category.Description FROM Category ORDER BY Category.Descrip"
                        "tion; "
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =120
                            Width =1410
                            Height =320
                            Name ="cmbDescription_Label"
                            Caption ="Description"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =247
                    OldBorderStyle =1
                    SpecialEffect =4
                    BorderWidth =3
                    Left =420
                    Top =1860
                    Width =7140
                    Height =6060
                    TabIndex =4
                    Name ="SubCategory"
                    SourceObject ="Form.CategorySub"
                    LinkChildFields ="CategoryID"
                    LinkMasterFields ="ID"

                    LayoutCachedLeft =420
                    LayoutCachedTop =1860
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =7920
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =0
                            Left =420
                            Top =1620
                            Width =1170
                            Height =315
                            Name ="Label10"
                            Caption ="Sub-Category"
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =255
                    Left =60
                    Top =60
                    Width =9000
                    Height =8580
                    Name ="Box11"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =8640
                End
                Begin ComboBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1920
                    Top =1260
                    Width =4560
                    Height =300
                    TabIndex =3
                    Name ="Combo14"
                    ControlSource ="SalesDepartmentID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SalesDepartment.ID, SalesDepartment.Description FROM SalesDepartment ORDE"
                        "R BY [ID]; "
                    ColumnWidths ="0;1440"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =0
                            Left =300
                            Top =1260
                            Width =1560
                            Height =320
                            Name ="Sales Department_Label"
                            Caption ="Sales Department"
                            EventProcPrefix ="Sales_Department_Label"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =255
                    Left =7080
                    Top =2460
                    Width =1695
                    Height =1320
                    TabIndex =5
                    Name ="cmdStep1_SetupCategories"
                    Caption ="Step 1. Set Subcategory IDENTITY_INSERT Off"
                    OnClick ="[Event Procedure]"

                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =255
                    Left =7020
                    Top =5880
                    Width =1695
                    Height =1080
                    TabIndex =6
                    Name ="cmdStep3_SetupOLCCCategories"
                    Caption ="Step3. Add  Categories as per OLCC"
                    OnClick ="[Event Procedure]"

                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =255
                    Left =7080
                    Top =4425
                    Width =1695
                    Height =960
                    TabIndex =7
                    Name ="cmdStep2_SetupCategories"
                    Caption ="Step 2. Add Subcategory 0 the set ID_INS ON "
                    OnClick ="[Event Procedure]"

                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =675
                    Top =8100
                    Width =2325
                    Height =330
                    ColumnWidth =1035
                    TabIndex =8
                    BackColor =10092543
                    Name ="Setup"
                    ControlSource ="Setup"

                    LayoutCachedLeft =675
                    LayoutCachedTop =8100
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =8430
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =120
                            Top =8100
                            Width =555
                            Height =315
                            Name ="Setup_Label"
                            Caption ="Setup"
                            LayoutCachedLeft =120
                            LayoutCachedTop =8100
                            LayoutCachedWidth =675
                            LayoutCachedHeight =8415
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6015
                    Top =8160
                    Width =2325
                    Height =330
                    ColumnWidth =1035
                    TabIndex =9
                    BackColor =10092543
                    Name ="Change"
                    ControlSource ="Change"

                    LayoutCachedLeft =6015
                    LayoutCachedTop =8160
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =8490
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =5340
                            Top =8160
                            Width =675
                            Height =315
                            Name ="Change_Label"
                            Caption ="Change"
                            LayoutCachedLeft =5340
                            LayoutCachedTop =8160
                            LayoutCachedWidth =6015
                            LayoutCachedHeight =8475
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =960
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
                            FontSize =8
                            Name ="lblDeleteRecord"
                            Caption ="D&elete"
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
                            FontSize =8
                            Name ="lblSaveRecord"
                            Caption ="Sa&ve"
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
                            FontSize =8
                            Name ="lblUndo"
                            Caption ="U&ndo"
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
                            FontSize =8
                            Name ="Label1"
                            Caption ="Close For&m"
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
                    FontSize =8
                    TabIndex =4
                    ForeColor =0
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
                    FontName ="MS Sans Serif"
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
                            FontSize =8
                            Name ="Label79"
                            Caption ="&Add New"
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
                            FontSize =8
                            Name ="Label61"
                            Caption ="Print Preview"
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
                            FontSize =8
                            Name ="Label60"
                            Caption ="Datasheet"
                            LayoutCachedLeft =5760
                            LayoutCachedTop =660
                            LayoutCachedWidth =6675
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    Left =60
                    Width =8640
                    Height =960
                    Name ="Box81"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =960
                End
            End
        End
    End
End
CodeBehindForm
' See "Category.cls"
