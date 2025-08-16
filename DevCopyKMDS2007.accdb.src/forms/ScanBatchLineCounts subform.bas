Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6638
    DatasheetFontHeight =11
    ItemSuffix =24
    Left =1140
    Top =2490
    Right =8070
    Bottom =7785
    RecSrcDt = Begin
        0xefbd482f5cf6e340
    End
    RecordSource ="ScanBatchLineCounts"
    Caption ="ScanBatchLineCounts subform"
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
            TextFontCharSet =204
            FontSize =11
            BorderColor =-2147483609
            ForeColor =11830108
            FontName ="Calibri"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =-2147483609
        End
        Begin CommandButton
            TextFontCharSet =204
            FontSize =11
            FontWeight =400
            ForeColor =-2147483609
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =-2147483609
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin BoundObjectFrame
            SizeMode =3
            BorderLineStyle =0
            BackStyle =0
            LabelX =-1800
            BorderColor =-2147483609
        End
        Begin TextBox
            FELineBreak = NotDefault
            TextFontCharSet =204
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            TextFontCharSet =204
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            TextFontCharSet =204
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin CustomControl
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin ToggleButton
            TextFontCharSet =204
            FontSize =11
            FontWeight =400
            ForeColor =-2147483609
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =204
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Attachment
            BackStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =4800
            Height =3840
            BorderColor =-2147483609
            LabelX =-1800
        End
        Begin PageHeader
            DisplayWhen =1
            Height =1260
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2550
                    Top =30
                    Width =3660
                    Height =330
                    Name ="ID"
                    ControlSource ="ID"
                    GroupTable =2
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2550
                    LayoutCachedTop =30
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =360
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =30
                            Width =2460
                            Height =330
                            Name ="ID_Label"
                            Caption ="ID"
                            GroupTable =2
                            BottomPadding =38
                            LayoutCachedLeft =30
                            LayoutCachedTop =30
                            LayoutCachedWidth =2490
                            LayoutCachedHeight =360
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2580
                    Top =360
                    Width =3660
                    Height =330
                    TabIndex =1
                    Name ="ScanBatchHeaderID"
                    ControlSource ="ScanBatchHeaderID"
                    GroupTable =3
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2580
                    LayoutCachedTop =360
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =690
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =60
                            Top =360
                            Width =2460
                            Height =330
                            Name ="ScanBatchHeaderID_Label"
                            Caption ="ScanBatchHeaderID"
                            GroupTable =3
                            BottomPadding =38
                            LayoutCachedLeft =60
                            LayoutCachedTop =360
                            LayoutCachedWidth =2520
                            LayoutCachedHeight =690
                            LayoutGroup =3
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2550
                    Top =780
                    Width =3660
                    Height =330
                    TabIndex =2
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    GroupTable =4
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2550
                    LayoutCachedTop =780
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =1110
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =780
                            Width =2460
                            Height =330
                            Name ="ItemID_Label"
                            Caption ="ItemID"
                            GroupTable =4
                            BottomPadding =38
                            LayoutCachedLeft =30
                            LayoutCachedTop =780
                            LayoutCachedWidth =2490
                            LayoutCachedHeight =1110
                            LayoutGroup =4
                            GroupTable =4
                        End
                    End
                End
            End
        End
        Begin Section
            Height =5295
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2880
                    Top =360
                    Width =3660
                    Height =330
                    ColumnWidth =1530
                    TabIndex =1
                    Name ="InputMethodID"
                    ControlSource ="InputMethod"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2880
                    LayoutCachedTop =360
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =690
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =360
                            Top =360
                            Width =2460
                            Height =330
                            Name ="InputMethodID_Label"
                            Caption ="InputMethodID"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =690
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2880
                    Top =765
                    Width =3660
                    Height =330
                    ColumnWidth =1530
                    TabIndex =2
                    Name ="CashierID"
                    ControlSource ="CashierID"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2880
                    LayoutCachedTop =765
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1095
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
                            Left =360
                            Top =765
                            Width =2460
                            Height =330
                            Name ="CashierID_Label"
                            Caption ="CashierID"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =765
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1095
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2880
                    Top =1170
                    Width =3660
                    Height =330
                    ColumnWidth =1380
                    TabIndex =3
                    Name ="LocationId"
                    ControlSource ="Location"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2880
                    LayoutCachedTop =1170
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1500
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
                            Left =360
                            Top =1170
                            Width =2460
                            Height =330
                            Name ="LocationId_Label"
                            Caption ="LocationId"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =1170
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1500
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2880
                    Top =1575
                    Width =3660
                    Height =330
                    ColumnWidth =2670
                    TabIndex =4
                    Name ="ScanCode"
                    ControlSource ="ScanCode"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2880
                    LayoutCachedTop =1575
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1905
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
                            Left =360
                            Top =1575
                            Width =2460
                            Height =330
                            Name ="ScanCode_Label"
                            Caption ="ScanCode"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =1575
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1905
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2880
                    Top =1980
                    Width =3660
                    Height =330
                    ColumnWidth =2325
                    TabIndex =5
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2880
                    LayoutCachedTop =1980
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =2310
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =1980
                            Width =2460
                            Height =330
                            Name ="SetupDate_Label"
                            Caption ="SetupDate"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =1980
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =2310
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2880
                    Top =60
                    Width =3660
                    Height =330
                    Name ="CountQuantity"
                    ControlSource ="CountQuantity"
                    GroupTable =5
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =390
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =5
                    GroupTable =5
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =360
                            Top =60
                            Width =2460
                            Height =330
                            Name ="CountQuantity_Label"
                            Caption ="CountQuantity"
                            GroupTable =5
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =60
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =390
                            LayoutGroup =5
                            GroupTable =5
                        End
                    End
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =360
            Name ="PageFooterSection"
            AutoHeight =1
        End
    End
End
