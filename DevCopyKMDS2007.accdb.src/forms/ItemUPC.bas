Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =255
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3480
    DatasheetFontHeight =10
    ItemSuffix =12
    Left =10710
    Top =5040
    Right =13200
    Bottom =6750
    AfterDelConfirm ="[Event Procedure]"
    OrderBy ="ChangeDate DESC"
    RecSrcDt = Begin
        0x89e69e2cfb92e240
    End
    RecordSource ="UPCItem"
    Caption ="UPCItem"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
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
        Begin FormHeader
            Height =480
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    Left =1800
                    Top =180
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    StatusBarText ="Change this UPC Code"
                    ControlTipText ="Change this UPC Code"
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
                            OverlapFlags =85
                            Top =180
                            Width =570
                            Height =240
                            Name ="Label3"
                            Caption ="ItemID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1020
                    Top =240
                    Width =540
                    Height =228
                    TabIndex =1
                    Name ="txtLastUsed"
                    ControlSource ="=DateDiff(\"d\",[ChangeDate],Now())"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000000000004000000000000000300000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x3900300000000000
                    End

                    LayoutCachedLeft =1020
                    LayoutCachedTop =240
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =468
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =960
                            Width =960
                            Height =180
                            Name ="Label11"
                            Caption ="Age"
                            LayoutCachedLeft =960
                            LayoutCachedWidth =1920
                            LayoutCachedHeight =180
                        End
                    End
                End
            End
        End
        Begin Section
            Height =1560
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =87
                    Left =60
                    Top =480
                    Width =2490
                    ColumnWidth =2052
                    Name ="ID"
                    ControlSource ="ID"
                    OnDblClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x010000008c000000020000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000010000000000000002000000150000000100000000000000 ,
                        0xcf7b790000000000000000000000000000000000000000000000000000000000 ,
                        0x000000006d0065002e007400780074004c006100730074005500730065006400 ,
                        0x3e0031003000300000000000
                    End

                    LayoutCachedLeft =60
                    LayoutCachedTop =480
                    LayoutCachedWidth =2550
                    LayoutCachedHeight =720
                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =240
                            Width =1590
                            Height =240
                            Name ="Label1"
                            Caption ="Alternate UPC Codes"
                            LayoutCachedLeft =120
                            LayoutCachedTop =240
                            LayoutCachedWidth =1710
                            LayoutCachedHeight =480
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =600
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =1920
                    ColumnWidth =2175
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
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
                            OverlapFlags =85
                            Left =120
                            Width =840
                            Height =240
                            Name ="Label5"
                            Caption ="SetupDate"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2040
                    Top =300
                    ColumnWidth =2160
                    TabIndex =1
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    LayoutCachedLeft =2040
                    LayoutCachedTop =300
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =540
                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =300
                            Width =975
                            Height =240
                            Name ="Label7"
                            Caption ="ChangeDate"
                            LayoutCachedLeft =240
                            LayoutCachedTop =300
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =540
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ItemUPC.cls"
