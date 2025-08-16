Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7290
    DatasheetFontHeight =11
    ItemSuffix =23
    Left =795
    Top =1155
    Right =14055
    Bottom =8055
    OrderBy ="[qryOLCCItemsNotOnFileSub].[Code] DESC, [qryOLCCItemsNotOnFileSub].[Description]"
    RecSrcDt = Begin
        0xdfdf2a8b54aee540
    End
    RecordSource ="qryOLCCItemsNotOnFileSub"
    Caption ="OLCCItemsNotOnFileSub"
    DatasheetFontName ="Calibri"
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
            Height =541
            Name ="FormHeader"
        End
        Begin Section
            Height =4203
            BackColor =13233915
            Name ="Detail"
            AlternateBackColor =13233915
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1965
                    Top =345
                    Width =3660
                    Height =330
                    ColumnWidth =1335
                    ColumnOrder =0
                    Name ="Code"
                    ControlSource ="Code"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1965
                    LayoutCachedTop =345
                    LayoutCachedWidth =5625
                    LayoutCachedHeight =675
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =345
                            Width =1560
                            Height =330
                            Name ="Code_Label"
                            Caption ="Code"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =345
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =675
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1965
                    Top =750
                    Width =3660
                    Height =330
                    ColumnWidth =3195
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1965
                    LayoutCachedTop =750
                    LayoutCachedWidth =5625
                    LayoutCachedHeight =1080
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =750
                            Width =1560
                            Height =330
                            Name ="Description_Label"
                            Caption ="Description"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =750
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =1080
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1965
                    Top =1155
                    Width =3660
                    Height =330
                    ColumnWidth =1095
                    ColumnOrder =4
                    TabIndex =3
                    Name ="Price"
                    ControlSource ="Price"
                    Format ="$#,##0.00;($#,##0.00)"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1965
                    LayoutCachedTop =1155
                    LayoutCachedWidth =5625
                    LayoutCachedHeight =1485
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =1155
                            Width =1560
                            Height =330
                            Name ="Price_Label"
                            Caption ="Price"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =1155
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =1485
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1965
                    Top =1560
                    Width =3660
                    Height =330
                    ColumnWidth =1035
                    ColumnOrder =5
                    TabIndex =4
                    Name ="UnitsPerCase"
                    ControlSource ="UnitsPerCase"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1965
                    LayoutCachedTop =1560
                    LayoutCachedWidth =5625
                    LayoutCachedHeight =1890
                    DatasheetCaption ="U/Case"
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =1560
                            Width =1560
                            Height =330
                            Name ="UnitsPerCase_Label"
                            Caption ="UnitsPerCase"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =1560
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =1890
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1986
                    Top =2842
                    Height =315
                    ColumnWidth =795
                    ColumnOrder =3
                    TabIndex =5
                    Name ="cmbSize"
                    ControlSource ="SizeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Size].[ID], [Size].[Description] FROM [Size]; "
                    ColumnWidths ="0;1440"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =1986
                    LayoutCachedTop =2842
                    LayoutCachedWidth =3687
                    LayoutCachedHeight =3157
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =285
                            Top =2842
                            Width =480
                            Height =305
                            Name ="Size_Label"
                            Caption ="Size"
                            LayoutCachedLeft =285
                            LayoutCachedTop =2842
                            LayoutCachedWidth =765
                            LayoutCachedHeight =3147
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =2156
                    Top =3352
                    Height =315
                    ColumnWidth =885
                    ColumnOrder =7
                    TabIndex =6
                    Name ="StockStatus"
                    ControlSource ="StockStatus"
                    RowSourceType ="Value List"
                    RowSource ="\"C\";\"Regular Stock (CS)\";\"H\";\"Holiday\";\"R\";\"Regular Stock\";\"S\";\"S"
                        "PECIAL ORDER!\";\"T\";\"One Time Buy\""
                    ColumnWidths ="0;1441"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =2156
                    LayoutCachedTop =3352
                    LayoutCachedWidth =3857
                    LayoutCachedHeight =3667
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =455
                            Top =3352
                            Width =1485
                            Height =305
                            Name ="cmbStockStatus_Label"
                            Caption ="Stock Status"
                            LayoutCachedLeft =455
                            LayoutCachedTop =3352
                            LayoutCachedWidth =1940
                            LayoutCachedHeight =3657
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2043
                    Top =2445
                    Height =315
                    ColumnWidth =630
                    ColumnOrder =6
                    TabIndex =7
                    Name ="SStatus"
                    ControlSource ="StockStatus"

                    LayoutCachedLeft =2043
                    LayoutCachedTop =2445
                    LayoutCachedWidth =3744
                    LayoutCachedHeight =2760
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =2445
                            Width =1110
                            Height =300
                            Name ="Label18"
                            Caption ="SS"
                            LayoutCachedLeft =342
                            LayoutCachedTop =2445
                            LayoutCachedWidth =1452
                            LayoutCachedHeight =2745
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1989
                    Top =2016
                    Height =315
                    ColumnWidth =2010
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Category"
                    ControlSource ="Category"
                    Format ="@"

                    LayoutCachedLeft =1989
                    LayoutCachedTop =2016
                    LayoutCachedWidth =3690
                    LayoutCachedHeight =2331
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =288
                            Top =2016
                            Width =840
                            Height =300
                            Name ="Label19"
                            Caption ="Category:"
                            LayoutCachedLeft =288
                            LayoutCachedTop =2016
                            LayoutCachedWidth =1128
                            LayoutCachedHeight =2316
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1986
                    Top =3750
                    Height =315
                    ColumnWidth =780
                    TabIndex =8
                    Name ="UPCCodeKMDS"
                    ControlSource ="UPCCodeKMDS"

                    LayoutCachedLeft =1986
                    LayoutCachedTop =3750
                    LayoutCachedWidth =3687
                    LayoutCachedHeight =4065
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =285
                            Top =3750
                            Width =1110
                            Height =300
                            Name ="Label21"
                            Caption ="UPC Code"
                            LayoutCachedLeft =285
                            LayoutCachedTop =3750
                            LayoutCachedWidth =1395
                            LayoutCachedHeight =4050
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5589
                    Top =3888
                    Height =315
                    TabIndex =9
                    Name ="SubCategoryID"
                    ControlSource ="SubCategoryID"

                    LayoutCachedLeft =5589
                    LayoutCachedTop =3888
                    LayoutCachedWidth =7290
                    LayoutCachedHeight =4203
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3888
                            Top =3888
                            Width =1005
                            Height =300
                            Name ="Label22"
                            Caption ="CategoryID:"
                            LayoutCachedLeft =3888
                            LayoutCachedTop =3888
                            LayoutCachedWidth =4893
                            LayoutCachedHeight =4188
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =462
            BackColor =13233915
            Name ="FormFooter"
            AutoHeight =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1992
                    Top =147
                    Height =315
                    ColumnWidth =1485
                    ColumnOrder =8
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    Format ="Medium Date"

                    LayoutCachedLeft =1992
                    LayoutCachedTop =147
                    LayoutCachedWidth =3693
                    LayoutCachedHeight =462
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =288
                            Top =144
                            Width =1110
                            Height =300
                            Name ="Label16"
                            Caption ="Downloaded"
                            LayoutCachedLeft =288
                            LayoutCachedTop =144
                            LayoutCachedWidth =1398
                            LayoutCachedHeight =444
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "OLCCItemsNotOnFileSub.cls"
