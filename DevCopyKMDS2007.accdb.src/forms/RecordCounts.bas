Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7746
    DatasheetFontHeight =10
    ItemSuffix =35
    Left =3045
    Top =7170
    Right =12120
    Bottom =10545
    RecSrcDt = Begin
        0x87f3483c55e5e240
    End
    DatasheetFontName ="Arial"
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
        Begin Section
            Height =2820
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1920
                    Top =2580
                    Width =960
                    Name ="Item"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1440
                            Top =2580
                            Width =405
                            Height =240
                            Name ="Label1"
                            Caption ="Item"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4665
                    Top =2580
                    Width =960
                    TabIndex =1
                    Name ="UPCItem"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3885
                            Top =2580
                            Width =705
                            Height =240
                            Name ="Label3"
                            Caption ="UPCItem"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5940
                    Top =420
                    Width =1386
                    Height =576
                    TabIndex =2
                    Name ="Command4"
                    Caption ="Get Record Counts"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Print Table"

                    LayoutCachedLeft =5940
                    LayoutCachedTop =420
                    LayoutCachedWidth =7326
                    LayoutCachedHeight =996
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4665
                    Top =780
                    Width =960
                    TabIndex =3
                    Name ="InvoLine"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3885
                            Top =780
                            Width =705
                            Height =240
                            Name ="Label6"
                            Caption ="InvoLine"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1920
                    Top =60
                    Width =960
                    TabIndex =4
                    Name ="Agents"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =780
                            Top =60
                            Width =1065
                            Height =240
                            Name ="Label8"
                            Caption ="Agents"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1920
                    Top =420
                    Width =960
                    TabIndex =5
                    Name ="Audit"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =780
                            Top =420
                            Width =1065
                            Height =240
                            Name ="Label11"
                            Caption ="Audit"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1920
                    Top =780
                    Width =960
                    TabIndex =6
                    Name ="Category"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =780
                            Top =780
                            Width =1065
                            Height =240
                            Name ="Label13"
                            Caption ="Category"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1920
                    Top =1140
                    Width =960
                    TabIndex =7
                    Name ="Cashier"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =780
                            Top =1140
                            Width =1065
                            Height =240
                            Name ="Label15"
                            Caption ="Cashier"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1920
                    Top =1500
                    Width =960
                    TabIndex =8
                    Name ="Club"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =780
                            Top =1500
                            Width =1065
                            Height =240
                            Name ="Label17"
                            Caption ="Club"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1920
                    Top =1860
                    Width =960
                    TabIndex =9
                    Name ="ClubOrderGuides"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =300
                            Top =1860
                            Width =1545
                            Height =240
                            Name ="Label19"
                            Caption ="ClubOrderGuides"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1920
                    Top =2220
                    Width =960
                    TabIndex =10
                    Name ="CRObjects"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =300
                            Top =2220
                            Width =1545
                            Height =240
                            Name ="Label21"
                            Caption ="CRObjects"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4665
                    Top =60
                    Width =960
                    TabIndex =11
                    Name ="ItemMovement"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3435
                            Top =60
                            Width =1155
                            Height =240
                            Name ="Label23"
                            Caption ="ItemMovement"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4665
                    Top =420
                    Width =960
                    TabIndex =12
                    Name ="ItemSalePrice"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3435
                            Top =420
                            Width =1155
                            Height =240
                            Name ="Label25"
                            Caption ="ItemSalePrice"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4665
                    Top =1140
                    Width =960
                    TabIndex =13
                    Name ="LotteryScratch"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3450
                            Top =1140
                            Width =1140
                            Height =240
                            Name ="Label27"
                            Caption ="LotteryScratch"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4665
                    Top =1500
                    Width =960
                    TabIndex =14
                    Name ="LotteryScratchRecon"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3000
                            Top =1500
                            Width =1590
                            Height =240
                            Name ="Label29"
                            Caption ="LotteryScratchRecon"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4665
                    Top =1860
                    Width =960
                    TabIndex =15
                    Name ="OrdHead"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3000
                            Top =1860
                            Width =1590
                            Height =240
                            Name ="Label31"
                            Caption ="OrdHead"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4665
                    Top =2220
                    Width =960
                    TabIndex =16
                    Name ="OrderLine"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3000
                            Top =2220
                            Width =1590
                            Height =240
                            Name ="Label33"
                            Caption ="OrderLine"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5880
                    Top =1860
                    Width =1755
                    Height =480
                    TabIndex =17
                    Name ="cmdPurgeItems"
                    Caption ="Delete Obsolete Items"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5880
                    LayoutCachedTop =1860
                    LayoutCachedWidth =7635
                    LayoutCachedHeight =2340
                End
            End
        End
    End
End
CodeBehindForm
' See "RecordCounts.cls"
