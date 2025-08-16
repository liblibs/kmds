Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8640
    DatasheetFontHeight =11
    ItemSuffix =23
    Left =3345
    Top =2415
    Right =13455
    Bottom =5460
    AfterInsert ="[Event Procedure]"
    RecSrcDt = Begin
        0x66f1133d5626e640
    End
    RecordSource ="SELECT KitItem.ID, KitItem.KitID, KitItem.ItemID, KitItem.ItemID, KitItem.Price,"
        " KitItem.SetupDate, KitItem.ChangeDate, IIf([ItemId]>99900000000 And [itemid]<99"
        "999999999,[Price],0) AS Liquor FROM KitItem;"
    Caption ="KitISubForm"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnError ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =10
            BorderColor =5137573
            ForeColor =5137573
            FontName ="Franklin Gothic Book"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
            BorderColor =5137573
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
            BorderColor =5137573
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BorderColor =5137573
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontWeight =400
            ForeColor =3160910
            FontName ="Franklin Gothic Book"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =5137573
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =5137573
        End
        Begin BoundObjectFrame
            SizeMode =3
            BorderLineStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
            BorderColor =5137573
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =5137573
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            BorderColor =5137573
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =5137573
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =5137573
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            Width =4536
            Height =2835
            BorderColor =5137573
        End
        Begin CustomControl
            OldBorderStyle =1
            Width =4536
            Height =2835
            BorderColor =5137573
        End
        Begin ToggleButton
            Width =283
            Height =283
            FontWeight =400
            ForeColor =3160910
            FontName ="Franklin Gothic Book"
            BorderLineStyle =0
        End
        Begin Tab
            Width =5103
            Height =3402
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Attachment
            BackStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =1701
            Height =1701
            BorderColor =5137573
            LabelX =-1701
        End
        Begin FormHeader
            Height =2042
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1650
                    Top =30
                    Width =1155
                    Height =360
                    Name ="KitID"
                    ControlSource ="KitID"
                    GroupTable =2
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1650
                    LayoutCachedTop =30
                    LayoutCachedWidth =2805
                    LayoutCachedHeight =390
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =30
                            Top =30
                            Width =1560
                            Height =360
                            Name ="KitID_Label"
                            Caption ="KitID"
                            GroupTable =2
                            BottomPadding =38
                            LayoutCachedLeft =30
                            LayoutCachedTop =30
                            LayoutCachedWidth =1590
                            LayoutCachedHeight =390
                            LayoutGroup =1
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1650
                    Top =840
                    Width =1155
                    Height =330
                    TabIndex =3
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    GroupTable =3
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1650
                    LayoutCachedTop =840
                    LayoutCachedWidth =2805
                    LayoutCachedHeight =1170
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =840
                            Width =1560
                            Height =330
                            Name ="SetupDate_Label"
                            Caption ="SetupDate"
                            GroupTable =3
                            BottomPadding =38
                            LayoutCachedLeft =30
                            LayoutCachedTop =840
                            LayoutCachedWidth =1590
                            LayoutCachedHeight =1170
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1650
                    Top =30
                    Width =1155
                    Height =330
                    TabIndex =1
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    GroupTable =3
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1650
                    LayoutCachedTop =30
                    LayoutCachedWidth =2805
                    LayoutCachedHeight =360
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =30
                            Top =30
                            Width =1560
                            Height =330
                            Name ="ChangeDate_Label"
                            Caption ="ChangeDate"
                            GroupTable =3
                            BottomPadding =38
                            LayoutCachedLeft =30
                            LayoutCachedTop =30
                            LayoutCachedWidth =1590
                            LayoutCachedHeight =360
                            LayoutGroup =2
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1650
                    Top =435
                    Width =1155
                    Height =330
                    TabIndex =2
                    Name ="upsize_ts"
                    ControlSource ="upsize_ts"
                    GroupTable =3
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1650
                    LayoutCachedTop =435
                    LayoutCachedWidth =2805
                    LayoutCachedHeight =765
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =435
                            Width =1560
                            Height =330
                            Name ="upsize_ts_Label"
                            Caption ="upsize_ts"
                            GroupTable =3
                            BottomPadding =38
                            LayoutCachedLeft =30
                            LayoutCachedTop =435
                            LayoutCachedWidth =1590
                            LayoutCachedHeight =765
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1845
                    Top =1245
                    Width =1155
                    Height =345
                    ColumnWidth =705
                    TabIndex =4
                    Name ="ID"
                    ControlSource ="ID"
                    GroupTable =4
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1845
                    LayoutCachedTop =1245
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =1590
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =225
                            Top =1245
                            Width =1560
                            Height =345
                            Name ="ID_Label"
                            Caption ="ID"
                            GroupTable =4
                            BottomPadding =38
                            LayoutCachedLeft =225
                            LayoutCachedTop =1245
                            LayoutCachedWidth =1785
                            LayoutCachedHeight =1590
                            LayoutGroup =3
                            GroupTable =4
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3798
                    Top =226
                    Width =906
                    Height =315
                    TabIndex =5
                    Name ="Text12"
                    ControlSource ="KitID"

                    LayoutCachedLeft =3798
                    LayoutCachedTop =226
                    LayoutCachedWidth =4704
                    LayoutCachedHeight =541
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3118
                            Top =253
                            Width =510
                            Height =300
                            Name ="Label13"
                            Caption ="KitID:"
                            LayoutCachedLeft =3118
                            LayoutCachedTop =253
                            LayoutCachedWidth =3628
                            LayoutCachedHeight =553
                        End
                    End
                End
            End
        End
        Begin Section
            Height =1905
            BackColor =13233915
            Name ="Detail"
            AlternateBackColor =13233915
            Begin
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    AccessKey =68
                    ColumnCount =7
                    ListRows =25
                    ListWidth =7560
                    Left =1637
                    Top =680
                    Width =2160
                    Height =330
                    ColumnWidth =2520
                    ColumnOrder =1
                    FontSize =8
                    BoundColumn =2
                    BackColor =16777088
                    BorderColor =0
                    Name ="cmbItemByDesc"
                    ControlSource ="Expr1002"
                    RowSourceType ="Table/Query"
                    RowSource ="lkupFormItemByDescription"
                    ColumnWidths ="2882;720;720;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =68
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1637
                    LayoutCachedTop =680
                    LayoutCachedWidth =3797
                    LayoutCachedHeight =1010
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =623
                            Top =725
                            Width =945
                            Height =285
                            FontSize =8
                            BorderColor =0
                            ForeColor =16711680
                            Name ="Label14"
                            Caption ="&Description"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =623
                            LayoutCachedTop =725
                            LayoutCachedWidth =1568
                            LayoutCachedHeight =1010
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    ColumnCount =8
                    ListRows =25
                    ListWidth =9000
                    Left =500
                    Top =226
                    Width =1410
                    Height =330
                    ColumnWidth =1770
                    ColumnOrder =0
                    FontSize =8
                    TabIndex =1
                    BoundColumn =3
                    BackColor =16777088
                    BorderColor =0
                    Name ="cmbItemByUPC"
                    ControlSource ="Expr1002"
                    RowSourceType ="Table/Query"
                    RowSource ="lkupFormItemByUPC"
                    ColumnWidths ="1442;2882;720;720;720;720;720;360"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =500
                    LayoutCachedTop =226
                    LayoutCachedWidth =1910
                    LayoutCachedHeight =556
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =56
                            Top =271
                            Width =375
                            Height =285
                            FontSize =8
                            BorderColor =0
                            ForeColor =16711680
                            Name ="Label15"
                            Caption ="UPC or Code"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =56
                            LayoutCachedTop =271
                            LayoutCachedWidth =431
                            LayoutCachedHeight =556
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4762
                    Top =1133
                    Height =315
                    TabIndex =2
                    Name ="Price"
                    ControlSource ="Price"

                    LayoutCachedLeft =4762
                    LayoutCachedTop =1133
                    LayoutCachedWidth =6463
                    LayoutCachedHeight =1448
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3061
                            Top =1133
                            Width =525
                            Height =300
                            Name ="Label16"
                            Caption ="Price:"
                            LayoutCachedLeft =3061
                            LayoutCachedTop =1133
                            LayoutCachedWidth =3586
                            LayoutCachedHeight =1433
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4581
                    Top =1590
                    Height =315
                    TabIndex =3
                    Name ="Liquor"
                    ControlSource ="Liquor"

                    LayoutCachedLeft =4581
                    LayoutCachedTop =1590
                    LayoutCachedWidth =6282
                    LayoutCachedHeight =1905
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2880
                            Top =1590
                            Width =525
                            Height =300
                            Name ="Label20"
                            Caption ="Liquor"
                            LayoutCachedLeft =2880
                            LayoutCachedTop =1590
                            LayoutCachedWidth =3405
                            LayoutCachedHeight =1890
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =315
            BackColor =13233915
            Name ="FormFooter"
            AutoHeight =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4478
                    Height =315
                    Name ="txtSumPrice"
                    ControlSource ="=Sum([Price])"

                    LayoutCachedLeft =4478
                    LayoutCachedWidth =6179
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2777
                            Width =525
                            Height =300
                            Name ="Label18"
                            Caption ="Price:"
                            LayoutCachedLeft =2777
                            LayoutCachedWidth =3302
                            LayoutCachedHeight =300
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1701
                    Height =315
                    TabIndex =1
                    Name ="txtSumLiquor"
                    ControlSource ="=Sum([Liquor])"

                    LayoutCachedLeft =1701
                    LayoutCachedWidth =3402
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Width =525
                            Height =300
                            Name ="Label22"
                            Caption ="Price:"
                            LayoutCachedWidth =525
                            LayoutCachedHeight =300
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "KitISubForm.cls"
