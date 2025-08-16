Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =27
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6780
    DatasheetFontHeight =10
    ItemSuffix =25
    Left =5736
    Top =5808
    Right =14100
    Bottom =11328
    DatasheetBackColor =65535
    RecSrcDt = Begin
        0xf125f1f80ae5e240
    End
    RecordSource ="SELECT ClubOrderGuides.ClubID, ClubOrderGuides.OrderID, ClubOrderGuides.LineID, "
        "ClubOrderGuides.ItemID, ClubOrderGuides.Description, ClubOrderGuides.Par, ClubOr"
        "derGuides.QuotedPrice, Club.NonLiquor, ClubOrderGuides.LiquorItem, Club.Contact,"
        " Club.MaintainParLevels FROM ClubOrderGuides INNER JOIN Club ON ClubOrderGuides."
        "ClubID=Club.ID; "
    Caption ="ClubOrderGuides"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Trebuchet MS"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =65535
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin ComboBox
            SpecialEffect =2
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin FormHeader
            Height =720
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =720
                    Top =60
                    ColumnOrder =0
                    Name ="ClubID"
                    ControlSource ="ClubID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =60
                            Width =555
                            Height =240
                            Name ="Label1"
                            Caption ="ClubID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2880
                    Top =60
                    Width =600
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Order"
                    ControlSource ="OrderID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2280
                            Top =60
                            Width =495
                            Height =240
                            Name ="Label3"
                            Caption ="Order"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3660
                    Top =90
                    ColumnOrder =2
                    TabIndex =2
                    Name ="chkNonLiquor"
                    ControlSource ="NonLiquor"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3890
                            Top =60
                            Width =855
                            Height =240
                            Name ="Label22"
                            Caption ="NonLiquor"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3660
                    Top =480
                    ColumnOrder =3
                    TabIndex =3
                    Name ="chkMaintainParLevels"
                    ControlSource ="MaintainParLevels"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3890
                            Top =450
                            Width =855
                            Height =240
                            Name ="Label24"
                            Caption ="NonLiquor"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =2280
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2100
                    Top =120
                    ColumnWidth =-2
                    Name ="LineID"
                    ControlSource ="LineID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =120
                            Width =525
                            Height =240
                            Name ="Label5"
                            Caption ="Line #"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2160
                    Top =600
                    ColumnWidth =-2
                    TabIndex =1
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.ID, Item.Description FROM Item ORDER BY Item.ID; "

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =720
                            Top =600
                            Width =630
                            Height =240
                            Name ="Label15"
                            Caption ="Code #"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2220
                    Top =1080
                    Width =2760
                    ColumnWidth =-2
                    TabIndex =2
                    Name ="Combo16"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.ID, Item.Description FROM Item ORDER BY Item.ID; "
                    ColumnWidths ="0;2880"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =780
                            Top =1080
                            Width =930
                            Height =240
                            Name ="Label17"
                            Caption ="Description"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2220
                    Top =1500
                    ColumnWidth =684
                    TabIndex =3
                    Name ="Par"
                    ControlSource ="Par"
                    Format ="General Number"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =780
                            Top =1500
                            Width =375
                            Height =240
                            Name ="Label18"
                            Caption ="Par:"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2220
                    Top =1920
                    ColumnWidth =780
                    TabIndex =4
                    Name ="QuotedPrice"
                    ControlSource ="QuotedPrice"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =780
                            Top =1920
                            Width =1020
                            Height =240
                            Name ="Label19"
                            Caption ="Quote:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =360
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "ClubOrderGuides.cls"
