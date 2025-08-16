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
    Width =7500
    DatasheetFontHeight =11
    ItemSuffix =28
    Left =1800
    Top =3630
    Right =13635
    Bottom =6450
    RecSrcDt = Begin
        0xe5da294359f8e340
    End
    RecordSource ="SELECT ScanBatchLineCounts.ID, ScanBatchLineCounts.ScanBatchHeaderID, ScanBatchL"
        "ineCounts.ItemID, ScanBatchLineCounts.InputMethod, ScanBatchLineCounts.CashierID"
        ", ScanBatchLineCounts.Location, ScanBatchLineCounts.ScanCode, ScanBatchLineCount"
        "s.CountQuantity, ScanBatchLineCounts.OnShelf, ScanBatchLineCounts.SetupDate, Sca"
        "nBatchLineCounts.ChangeDate FROM ScanBatchLineCounts; "
    Caption ="ScanBatchLineCountsSub"
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
        Begin FormHeader
            Height =2805
            BackColor =-2147483609
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =85
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
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2550
                    Top =435
                    Width =3660
                    Height =330
                    TabIndex =1
                    Name ="ScanBatchHeaderID"
                    ControlSource ="ScanBatchHeaderID"
                    GroupTable =2
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2550
                    LayoutCachedTop =435
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =765
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =435
                            Width =2460
                            Height =330
                            Name ="ScanBatchHeaderID_Label"
                            Caption ="ScanBatchHeaderID"
                            GroupTable =2
                            BottomPadding =38
                            LayoutCachedLeft =30
                            LayoutCachedTop =435
                            LayoutCachedWidth =2490
                            LayoutCachedHeight =765
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2580
                    Top =2040
                    Width =3660
                    Height =330
                    TabIndex =2
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    GroupTable =4
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2580
                    LayoutCachedTop =2040
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =2370
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =60
                            Top =2040
                            Width =2460
                            Height =330
                            Name ="ChangeDate_Label"
                            Caption ="ChangeDate"
                            GroupTable =4
                            BottomPadding =38
                            LayoutCachedLeft =60
                            LayoutCachedTop =2040
                            LayoutCachedWidth =2520
                            LayoutCachedHeight =2370
                            LayoutGroup =3
                            GroupTable =4
                        End
                    End
                End
            End
        End
        Begin Section
            Height =4890
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3000
                    Top =360
                    Width =3660
                    Height =330
                    ColumnWidth =1050
                    ColumnOrder =0
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =3000
                    LayoutCachedTop =360
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =690
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =480
                            Top =360
                            Width =2460
                            Height =330
                            Name ="ItemID_Label"
                            Caption ="ItemID"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =480
                            LayoutCachedTop =360
                            LayoutCachedWidth =2940
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
                    Left =3000
                    Top =765
                    Width =3660
                    Height =330
                    ColumnWidth =1755
                    ColumnOrder =1
                    TabIndex =2
                    Name ="ScanCode"
                    ControlSource ="ScanCode"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =3000
                    LayoutCachedTop =765
                    LayoutCachedWidth =6660
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
                            Left =480
                            Top =765
                            Width =2460
                            Height =330
                            Name ="ScanCode_Label"
                            Caption ="ScanCode"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =480
                            LayoutCachedTop =765
                            LayoutCachedWidth =2940
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
                    Left =3000
                    Top =1170
                    Width =3660
                    Height =330
                    ColumnWidth =1530
                    TabIndex =3
                    Name ="CountQuantity"
                    ControlSource ="CountQuantity"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =3000
                    LayoutCachedTop =1170
                    LayoutCachedWidth =6660
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
                            Left =480
                            Top =1170
                            Width =2460
                            Height =330
                            Name ="CountQuantity_Label"
                            Caption ="CountQuantity"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =480
                            LayoutCachedTop =1170
                            LayoutCachedWidth =2940
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
                    Left =3000
                    Top =1575
                    Width =3660
                    Height =330
                    ColumnWidth =2235
                    TabIndex =4
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =3000
                    LayoutCachedTop =1575
                    LayoutCachedWidth =6660
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
                            Left =480
                            Top =1575
                            Width =2460
                            Height =330
                            Name ="SetupDate_Label"
                            Caption ="SetupDate"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =480
                            LayoutCachedTop =1575
                            LayoutCachedWidth =2940
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
                    Left =3000
                    Top =2160
                    Width =3660
                    Height =330
                    ColumnOrder =2
                    TabIndex =1
                    Name ="LocationId"
                    ControlSource ="Location"
                    GroupTable =5
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =3000
                    LayoutCachedTop =2160
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =2490
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GroupTable =5
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =480
                            Top =2160
                            Width =2460
                            Height =330
                            Name ="LocationId_Label"
                            Caption ="LocationId"
                            GroupTable =5
                            BottomPadding =38
                            LayoutCachedLeft =480
                            LayoutCachedTop =2160
                            LayoutCachedWidth =2940
                            LayoutCachedHeight =2490
                            LayoutGroup =4
                            GroupTable =5
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3000
                    Top =2580
                    Width =3660
                    Height =330
                    TabIndex =5
                    Name ="InputMethodID"
                    ControlSource ="InputMethod"
                    GroupTable =6
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =3000
                    LayoutCachedTop =2580
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =2910
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =5
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =480
                            Top =2580
                            Width =2460
                            Height =330
                            Name ="InputMethodID_Label"
                            Caption ="InputMethodID"
                            GroupTable =6
                            BottomPadding =38
                            LayoutCachedLeft =480
                            LayoutCachedTop =2580
                            LayoutCachedWidth =2940
                            LayoutCachedHeight =2910
                            LayoutGroup =5
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3000
                    Top =3000
                    Width =3660
                    Height =330
                    ColumnWidth =630
                    TabIndex =6
                    Name ="CashierID"
                    ControlSource ="CashierID"
                    GroupTable =7
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =3000
                    LayoutCachedTop =3000
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =3330
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =6
                    GroupTable =7
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =480
                            Top =3000
                            Width =2460
                            Height =330
                            Name ="CashierID_Label"
                            Caption ="CashierID"
                            GroupTable =7
                            BottomPadding =38
                            LayoutCachedLeft =480
                            LayoutCachedTop =3000
                            LayoutCachedWidth =2940
                            LayoutCachedHeight =3330
                            LayoutGroup =6
                            GroupTable =7
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2940
                    Top =3360
                    Width =3660
                    Height =330
                    ColumnWidth =1170
                    TabIndex =7
                    Name ="OnShelf"
                    ControlSource ="OnShelf"
                    GroupTable =8
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2940
                    LayoutCachedTop =3360
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =3690
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =7
                    GroupTable =8
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =420
                            Top =3360
                            Width =2460
                            Height =330
                            Name ="OnShelf_Label"
                            Caption ="OnShelf"
                            GroupTable =8
                            BottomPadding =38
                            LayoutCachedLeft =420
                            LayoutCachedTop =3360
                            LayoutCachedWidth =2880
                            LayoutCachedHeight =3690
                            LayoutGroup =7
                            GroupTable =8
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =848
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2550
                    Top =30
                    Width =3660
                    Height =330
                    Name ="txtTotalSum"
                    ControlSource ="=Sum([CountQuantity])"
                    GroupTable =10
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2550
                    LayoutCachedTop =30
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =360
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =8
                    GroupTable =10
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =30
                            Width =2460
                            Height =330
                            Name ="Label25"
                            Caption ="CountQuantity"
                            GroupTable =10
                            BottomPadding =38
                            LayoutCachedLeft =30
                            LayoutCachedTop =30
                            LayoutCachedWidth =2490
                            LayoutCachedHeight =360
                            LayoutGroup =8
                            GroupTable =10
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2550
                    Top =480
                    Width =3660
                    Height =330
                    TabIndex =1
                    Name ="txtTotalCount"
                    ControlSource ="=Count([CountQuantity])"
                    GroupTable =11
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2550
                    LayoutCachedTop =480
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =810
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =9
                    GroupTable =11
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =480
                            Width =2460
                            Height =330
                            Name ="Label27"
                            Caption ="CountQuantity"
                            GroupTable =11
                            BottomPadding =38
                            LayoutCachedLeft =30
                            LayoutCachedTop =480
                            LayoutCachedWidth =2490
                            LayoutCachedHeight =810
                            LayoutGroup =9
                            GroupTable =11
                        End
                    End
                End
            End
        End
    End
End
