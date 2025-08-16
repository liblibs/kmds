Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =255
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10980
    DatasheetFontHeight =10
    ItemSuffix =71
    Left =5055
    Right =16125
    Bottom =5235
    HelpContextId =4200
    OnUnload ="[Event Procedure]"
    ShortcutMenuBar ="KMDSPopup"
    RecSrcDt = Begin
        0x5f1feb0e6a94e240
    End
    Caption ="Onhands - Cigarettes - Premium"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnResize ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin FormHeader
            Visible = NotDefault
            CanGrow = NotDefault
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin PageHeader
            DisplayWhen =1
            Height =660
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =10620
                    Height =420
                    FontSize =16
                    Name ="txtReportTitle"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =60
                            Width =645
                            Height =240
                            Name ="Label53"
                            Caption ="Text52:"
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =3240
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Top =60
                    Width =10800
                    Height =3120
                    Name ="ItemDataSheetSub"
                    SourceObject ="Form.ItemDataSheetSub"

                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =1200
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Top =900
                    Width =8520
                    Height =300
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text48"
                    ControlSource ="=Now()"
                    Format ="Long Time"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    SpecialEffect =0
                    OverlapFlags =255
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =0
                    Left =5640
                    Top =900
                    Width =5040
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    SpecialEffect =0
                    OverlapFlags =127
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Top =600
                    Width =8520
                    Height =300
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text63"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    SpecialEffect =0
                    OverlapFlags =255
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =0
                    Left =5640
                    Top =600
                    Width =5040
                    Height =300
                    FontWeight =700
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text64"
                    ControlSource ="=[Name]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Top =900
                    Width =10680
                    Height =300
                    FontSize =6
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="Text65"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Top =600
                    Width =10680
                    Height =255
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =1995
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1200
                    Top =120
                    Width =1080
                    Name ="Text55"
                    ControlSource ="=ItemDataSheetSub.Form!txtTotalOnhand"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =120
                            Width =1050
                            Height =240
                            Name ="Label56"
                            Caption ="Total Onhand"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2820
                    Top =120
                    Width =1260
                    TabIndex =1
                    Name ="Text57"
                    ControlSource ="=ItemDataSheetSub.Form!txtTotalRetail"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =2820
                    LayoutCachedTop =120
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2340
                            Top =120
                            Width =405
                            Height =240
                            Name ="Label58"
                            Caption =" Retail"
                            LayoutCachedLeft =2340
                            LayoutCachedTop =120
                            LayoutCachedWidth =2745
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4605
                    Top =120
                    Width =1080
                    TabIndex =2
                    Name ="Text59"
                    ControlSource ="=ItemDataSheetSub.Form!txtTotalCost"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =4605
                    LayoutCachedTop =120
                    LayoutCachedWidth =5685
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =4080
                            Top =120
                            Width =450
                            Height =240
                            Name ="Label60"
                            Caption =" Cost"
                            LayoutCachedLeft =4080
                            LayoutCachedTop =120
                            LayoutCachedWidth =4530
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6150
                    Top =120
                    Width =960
                    TabIndex =3
                    Name ="Text61"
                    ControlSource ="=ItemDataSheetSub.Form!txtTotalRetail-ItemDataSheetSub.Form!txtTotalCost"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =6150
                    LayoutCachedTop =120
                    LayoutCachedWidth =7110
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5760
                            Top =120
                            Width =315
                            Height =240
                            Name ="Label62"
                            Caption =" GP"
                            LayoutCachedLeft =5760
                            LayoutCachedTop =120
                            LayoutCachedWidth =6075
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7800
                    Top =120
                    Width =660
                    TabIndex =4
                    Name ="Text66"
                    ControlSource ="=((ItemDataSheetSub.Form!txtTotalRetail-ItemDataSheetSub.Form!txtTotalCost)/Item"
                        "DataSheetSub.Form!txtTotalRetail)"
                    Format ="Percent"

                    LayoutCachedLeft =7800
                    LayoutCachedTop =120
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7200
                            Top =120
                            Width =525
                            Height =240
                            Name ="Label67"
                            Caption =" GP%"
                            LayoutCachedLeft =7200
                            LayoutCachedTop =120
                            LayoutCachedWidth =7725
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =67
                    Left =9660
                    Top =60
                    Width =606
                    Height =426
                    TabIndex =5
                    ForeColor =0
                    Name ="cmdfrmClose"
                    Caption ="Command64"
                    OnClick ="[Event Procedure]"
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
                    UnicodeAccessKey =67

                    LayoutCachedLeft =9660
                    LayoutCachedTop =60
                    LayoutCachedWidth =10266
                    LayoutCachedHeight =486
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =10320
                            Top =120
                            Width =555
                            Height =285
                            Name ="Label65"
                            Caption ="&Close"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =10320
                            LayoutCachedTop =120
                            LayoutCachedWidth =10875
                            LayoutCachedHeight =405
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Top =540
                    Width =10860
                    Height =1275
                    TabIndex =6
                    Name ="qryItemMovementByMonth"
                    SourceObject ="Form.PurchaseOrdersOnFileSubItemMovement"

                    LayoutCachedTop =540
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =1815
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =69
                    Left =8580
                    Top =60
                    Width =606
                    Height =426
                    TabIndex =7
                    ForeColor =0
                    Name ="cmdOutput"
                    Caption ="Command64"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"
                    UnicodeAccessKey =69
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
                        0x0000000000000000000000000000000000000000208030ff207820ff207820ff ,
                        0x207820ff107020ff107020ff107010ff106810ff106810ff106810ff00000000 ,
                        0x0000000000000000000000000000000000000000208030ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffff106810ff00000000 ,
                        0x0000000000000000000000000000000000000000308030ffffffffff205810ff ,
                        0x205810ff205810ff90e090ff309840ff308030ffffffffff106810ff00000000 ,
                        0x0000000000000000000000000000000000000000308830ffffffffffa0d8a0ff ,
                        0x205810ffa0e0a0ff30a060ff40a870ff90e080ffffffffff107010ff00000000 ,
                        0x0000000000000000000000000000000000000000308830ffffffffffc0e8c0ff ,
                        0x90c090ff205810ff40b080ffa0e090ff90e090ffffffffff107020ff00000000 ,
                        0x0000000000000000000000000000000000000000308830fff0f8f0ffd0f0d0ff ,
                        0xb0e0b0ff40a870ff40a870ffa0e0a0ffa0e0a0ffffffffff107020ff00000000 ,
                        0x0000000000000000000000000000000000000000308830ffe0f0e0ffe0f8e0ff ,
                        0x309040ff409860ff205810ff205810ffa0d8a0ffffffffff207820ff00000000 ,
                        0x0000000000000000000000000000000000000000308830d0c0e0c0ff308020ff ,
                        0x308030ffd0f0d0ffb0d0b0ff205810ff608050ffffffffff207820ff00000000 ,
                        0x00000000000000000000000000000000000000003088308080c090ffc0e0c0ff ,
                        0xe0f0e0fff0f8f0ffffffffffffffffffffffffffffffffff207820ff00000000 ,
                        0x00000000000000000000000000000000000000000000000030883080308830d0 ,
                        0x308830ff308830ff308830ff308830ff308030ff208030ff208030ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =8580
                    LayoutCachedTop =60
                    LayoutCachedWidth =9186
                    LayoutCachedHeight =486
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =9240
                            Top =120
                            Width =555
                            Height =285
                            Name ="Label70"
                            Caption ="&Export"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =9240
                            LayoutCachedTop =120
                            LayoutCachedWidth =9795
                            LayoutCachedHeight =405
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ItemDataSheet.cls"
