Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11340
    DatasheetFontHeight =10
    ItemSuffix =161
    Top =90
    OnNoData ="[Event Procedure]"
    OnPage ="[Event Procedure]"
    OrderBy ="LineID"
    RecSrcDt = Begin
        0x026e791efa14e440
    End
    RecordSource ="qryClubInvoiceAndSysfile"
    Caption ="Club Pick List"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Trebuchet MS"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin ListBox
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="ID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="CatDesc"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1200
            OnFormat ="[Event Procedure]"
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =8700
                    Top =60
                    Width =1980
                    Height =540
                    FontSize =20
                    Name ="Label50"
                    Caption ="Pick List"
                    FontName ="Trebuchet MS"
                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =4440
                    Height =255
                    ColumnOrder =0
                    Name ="Text66"
                    ControlSource ="ClubContact"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =300
                    Width =4440
                    Height =255
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text67"
                    ControlSource ="ClubAddress1"

                    LayoutCachedLeft =60
                    LayoutCachedTop =300
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =555
                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =540
                    Width =4440
                    Height =255
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text68"
                    ControlSource ="ClubAddress2"

                    LayoutCachedLeft =60
                    LayoutCachedTop =540
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =795
                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =780
                    Width =4440
                    Height =255
                    ColumnOrder =3
                    TabIndex =3
                    Name ="Text69"
                    ControlSource ="ClubCityST"

                    LayoutCachedLeft =60
                    LayoutCachedTop =780
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =1035
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =10440
                    Top =120
                    Width =720
                    ColumnOrder =4
                    TabIndex =4
                    Name ="txtPages"
                    ControlSource ="=[Pages]"

                    LayoutCachedLeft =10440
                    LayoutCachedTop =120
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =8715
                    Top =600
                    Width =1965
                    Height =540
                    ColumnOrder =5
                    FontSize =20
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="Text70"
                    ControlSource ="InvoiceNo"
                    FontName ="Trebuchet MS"

                End
                Begin CheckBox
                    Visible = NotDefault
                    Left =10800
                    Top =540
                    ColumnOrder =6
                    TabIndex =6
                    Name ="NonLiquor"
                    ControlSource ="NonLiquor"

                    LayoutCachedLeft =10800
                    LayoutCachedTop =540
                    LayoutCachedWidth =11060
                    LayoutCachedHeight =780
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =1
                    Left =4560
                    Top =60
                    Width =4081
                    Height =1080
                    TabIndex =7
                    Name ="qryClubOrderSizeTotals"
                    SourceObject ="Report.ClubOrderSizeTotals"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =60
                    LayoutCachedWidth =8641
                    LayoutCachedHeight =1140
                End
            End
        End
        Begin PageHeader
            Height =720
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9120
                    Top =360
                    Width =1545
                    Height =300
                    Name ="Text212"
                    ControlSource ="=[Page] & \" of \" & [Pages]"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =9120
                            Top =60
                            Width =1545
                            Height =300
                            BackColor =13882323
                            Name ="Label213"
                            Caption ="Page Number"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =120
                    Top =60
                    Width =8520
                    Height =540
                    FontSize =20
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text156"
                    ControlSource ="ClubName"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =600
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            RepeatSection = NotDefault
            Height =420
            Name ="GroupHeader0"
            Begin
                Begin Label
                    TextAlign =2
                    Left =1020
                    Top =60
                    Width =900
                    Height =300
                    Name ="Label83"
                    Caption ="Item"
                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =360
                End
                Begin Label
                    TextAlign =2
                    Left =2400
                    Top =60
                    Width =2640
                    Height =300
                    Name ="Label84"
                    Caption ="Description"
                    LayoutCachedLeft =2400
                    LayoutCachedTop =60
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =360
                End
                Begin Label
                    TextAlign =2
                    Left =180
                    Top =60
                    Width =555
                    Height =300
                    Name ="Label85"
                    Caption ="Qty"
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =735
                    LayoutCachedHeight =360
                End
                Begin Label
                    TextAlign =2
                    Left =6720
                    Top =60
                    Width =720
                    Height =300
                    Name ="Label177"
                    Caption ="Size"
                    LayoutCachedLeft =6720
                    LayoutCachedTop =60
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =8520
                    Width =0
                    Height =420
                    Name ="Line155"
                    LayoutCachedLeft =8520
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =2
                    Left =7590
                    Top =60
                    Width =885
                    Height =300
                    Name ="Label153"
                    Caption ="On Shelf"
                    LayoutCachedLeft =7590
                    LayoutCachedTop =60
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =60
                    Width =0
                    Height =420
                    Name ="Line158"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =60
                    LayoutCachedHeight =420
                End
                Begin Line
                    Left =60
                    Width =8460
                    Name ="Line159"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =8520
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =375
            BreakLevel =1
            Name ="GroupHeader1"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =4695
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="CatDesc"
                    ControlSource ="CatDesc"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =4815
                    LayoutCachedHeight =345
                End
                Begin Line
                    Left =8520
                    Width =0
                    Height =360
                    Name ="Line146"
                    LayoutCachedLeft =8520
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =60
                    Width =0
                    Height =360
                    Name ="Line150"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =60
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =60
                    Top =360
                    Width =8460
                    Name ="Line151"
                    LayoutCachedLeft =60
                    LayoutCachedTop =360
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =60
                    Width =8460
                    Name ="Line152"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =8520
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =420
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Line
                    Left =8520
                    Width =0
                    Height =420
                    Name ="Line154"
                    LayoutCachedLeft =8520
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =420
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    BackStyle =1
                    IMESentenceMode =3
                    ColumnCount =15
                    ListRows =1
                    Left =7620
                    Top =60
                    Width =720
                    Height =360
                    FontSize =11
                    BoundColumn =12
                    Name ="cmbShelf"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT * FROM lkupCluOrderItemByCode; "
                    ColumnWidths ="0;0;0;1080;0;0;0;0;0;0;0;0;0;0;0;0"
                    ShortcutMenuBar ="KMDSPopup"

                    LayoutCachedLeft =7620
                    LayoutCachedTop =60
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =420
                End
                Begin Line
                    Left =6540
                    Width =0
                    Height =420
                    Name ="Line135"
                    LayoutCachedLeft =6540
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =420
                End
                Begin Line
                    Left =2220
                    Width =0
                    Height =420
                    Name ="Line134"
                    LayoutCachedLeft =2220
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    BackStyle =1
                    Left =6780
                    Top =60
                    Width =600
                    Height =360
                    FontSize =11
                    TabIndex =1
                    Name ="Text129"
                    ControlSource ="SizeDesc"

                    LayoutCachedLeft =6780
                    LayoutCachedTop =60
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =420
                End
                Begin Line
                    Left =720
                    Width =0
                    Height =420
                    Name ="Line93"
                    LayoutCachedLeft =720
                    LayoutCachedWidth =720
                    LayoutCachedHeight =420
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =60
                    Width =0
                    Height =420
                    Name ="Line92"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =60
                    LayoutCachedHeight =420
                End
                Begin Line
                    Left =7560
                    Width =0
                    Height =420
                    Name ="Line91"
                    LayoutCachedLeft =7560
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =120
                    Top =60
                    Width =540
                    Height =360
                    FontSize =11
                    TabIndex =2
                    Name ="QuantityOrdered"
                    ControlSource ="QuantityShipped"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =660
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    BackStyle =1
                    Left =2280
                    Top =60
                    Width =4200
                    Height =360
                    FontSize =12
                    TabIndex =3
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2280
                    LayoutCachedTop =60
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =780
                    Top =60
                    Width =1365
                    Height =360
                    FontSize =10
                    TabIndex =4
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =780
                    LayoutCachedTop =60
                    LayoutCachedWidth =2145
                    LayoutCachedHeight =420
                End
            End
        End
        Begin PageFooter
            Visible = NotDefault
            Height =1020
            OnFormat ="[Event Procedure]"
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Width =4320
                    Name ="txtDate"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7200
                    Top =240
                    Width =3480
                    TabIndex =1
                    Name ="txtCashier"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Top =240
                    Width =4320
                    TabIndex =2
                    Name ="Text217"
                    ControlSource ="=Now()"
                    Format ="Long Time"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7200
                    Width =3480
                    TabIndex =3
                    Name ="Text218"
                    ControlSource ="=[Name]"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Top =300
                    Width =10680
                    Height =180
                    FontSize =6
                    TabIndex =4
                    Name ="Text219"
                    ControlSource ="=\"Kelly Myers Data Systems\""

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Width =10680
                    TabIndex =5
                    Name ="txtStoreName"
                    Format ="Long Date"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =15
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin Line
                    Left =60
                    Width =8460
                    Name ="Line109"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =8520
                End
            End
        End
    End
End
CodeBehindForm
' See "ClubPickListByCategory.cls"
