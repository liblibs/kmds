Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8340
    DatasheetFontHeight =11
    ItemSuffix =11
    Left =5130
    Top =555
    Right =13470
    Bottom =8295
    RecSrcDt = Begin
        0xc7ec3f1d0ff9e340
    End
    RecordSource ="qryScanBatchLineCountsTotal"
    Caption ="ScanBatchTotalsSub"
    OnCurrent ="[Event Procedure]"
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
            Visible = NotDefault
            Height =1380
            BackColor =-2147483609
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2550
                    Top =30
                    Width =2400
                    Height =360
                    Name ="ScanBatchHeaderID"
                    ControlSource ="ScanBatchHeaderID"
                    GroupTable =2
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2550
                    LayoutCachedTop =30
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =390
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =30
                            Top =30
                            Width =2460
                            Height =360
                            Name ="ScanBatchHeaderID_Label"
                            Caption ="ScanBatchHeaderID"
                            GroupTable =2
                            BottomPadding =38
                            LayoutCachedLeft =30
                            LayoutCachedTop =30
                            LayoutCachedWidth =2490
                            LayoutCachedHeight =390
                            LayoutGroup =2
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2580
                    Top =360
                    Width =2400
                    Height =330
                    TabIndex =1
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    GroupTable =3
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2580
                    LayoutCachedTop =360
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =690
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =60
                            Top =360
                            Width =2460
                            Height =330
                            Name ="ItemID_Label"
                            Caption ="ItemID"
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7725
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2880
                    Top =810
                    Width =2400
                    Height =330
                    ColumnWidth =1050
                    TabIndex =2
                    Name ="Count"
                    ControlSource ="SumOfCountQuantity"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2880
                    LayoutCachedTop =810
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =1140
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =360
                            Top =810
                            Width =2460
                            Height =330
                            Name ="SumOfCountQuantity_Label"
                            Caption ="Total Counted"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =810
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1140
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
                    Top =420
                    Width =2400
                    Height =315
                    Name ="OnShelf"
                    ControlSource ="LastOfOnShelf"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2880
                    LayoutCachedTop =420
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =735
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =360
                            Top =420
                            Width =2460
                            Height =315
                            Name ="Label6"
                            Caption ="On Shelf:"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =360
                            LayoutCachedTop =420
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =735
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =360
                    Top =1380
                    Width =7200
                    Height =5940
                    TabIndex =1
                    Name ="ScanBatchLineCounts subform"
                    SourceObject ="Form.ScanBatchLineCounts subform"
                    LinkChildFields ="ScanBatchHeaderID;ItemID"
                    LinkMasterFields ="ScanBatchHeaderID;ItemID"
                    EventProcPrefix ="ScanBatchLineCounts_subform"

                    LayoutCachedLeft =360
                    LayoutCachedTop =1380
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =7320
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =2880
                    Height =315
                    TabIndex =3
                    Name ="cmbItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [qryScanBatchLineCountsTotal].[ItemID] FROM qryScanBatchLineCountsTotal; "
                    ColumnWidths ="1440"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =2880
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =360
                            Width =2460
                            Height =320
                            Name ="ItemCode_Label"
                            Caption ="Item Code"
                            LayoutCachedLeft =360
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =320
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
        End
    End
End
CodeBehindForm
' See "ScanBatchTotalsSub.cls"
