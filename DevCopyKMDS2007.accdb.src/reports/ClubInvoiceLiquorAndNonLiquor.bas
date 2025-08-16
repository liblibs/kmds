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
    Width =10695
    DatasheetFontHeight =10
    ItemSuffix =185
    Top =285
    OnNoData ="[Event Procedure]"
    OnPage ="[Event Procedure]"
    OrderBy ="ItemID"
    RecSrcDt = Begin
        0x8a6745938f42e340
    End
    RecordSource ="qryClubInvoiceLiquorAndNonLiquor"
    Caption ="ClubInvoiceLiquorAndNonLiquor"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Trebuchet MS"
    FilterOnLoad =255
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
            GroupFooter = NotDefault
            ControlSource ="LiquorItem"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2160
            OnFormat ="[Event Procedure]"
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =1
                    BackStyle =1
                    Left =60
                    Top =360
                    Width =7320
                    Height =540
                    ColumnOrder =0
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text56"
                    ControlSource ="ClubName"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    Left =7260
                    Top =360
                    Width =3360
                    Height =540
                    ColumnOrder =1
                    FontSize =20
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text70"
                    ControlSource ="=\"Invoice #\" & [InvoiceNo]"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =9000
                    Top =1260
                    Width =1620
                    Height =420
                    ColumnOrder =2
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="Text73"
                    ControlSource ="ExtNetPrice"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =9000
                    Top =960
                    Width =1620
                    Height =300
                    BackColor =16053492
                    Name ="lblInvoiceAmount"
                    Caption ="Invoice Amount"
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =180
                    Top =1260
                    Width =1635
                    Height =420
                    ColumnOrder =3
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    Name ="Text140"
                    ControlSource ="InvDate"
                    Format ="Short Date"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =180
                            Top =960
                            Width =1635
                            Height =300
                            BackColor =16053492
                            Name ="Label141"
                            Caption ="Invoice Date"
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =2940
                    Top =1380
                    Width =4680
                    Height =300
                    BackColor =16053492
                    Name ="Label176"
                    Caption ="Combined Invoice"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =180
                    Top =1740
                    Width =10440
                    Height =420
                    ColumnOrder =4
                    TabIndex =4
                    Name ="txtOrderNotes"
                    ControlSource ="OrderNotes"

                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =2
                    Width =10620
                    Height =420
                    FontSize =16
                    Name ="Label169"
                    Caption ="Customer Copy"
                End
            End
        End
        Begin PageHeader
            Height =2580
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9060
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
                            Left =9060
                            Top =60
                            Width =1545
                            Height =300
                            BackColor =16053492
                            Name ="Label213"
                            Caption ="Page Number"
                        End
                    End
                End
                Begin TextBox
                    Left =4920
                    Top =420
                    Width =3420
                    Height =270
                    TabIndex =1
                    Name ="Text77"
                    ControlSource ="ClubName"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =4800
                            Top =60
                            Width =3720
                            Height =300
                            BackColor =16053492
                            Name ="Label69"
                            Caption ="    Ship to / Sold to"
                        End
                    End
                End
                Begin TextBox
                    Left =4920
                    Top =660
                    Width =3420
                    Height =270
                    TabIndex =2
                    Name ="Text78"
                    ControlSource ="ClubContact"

                End
                Begin TextBox
                    Left =4920
                    Top =900
                    Width =3420
                    Height =270
                    TabIndex =3
                    Name ="Text80"
                    ControlSource ="ClubAddress1"

                End
                Begin TextBox
                    Left =4920
                    Top =1200
                    Width =3420
                    Height =270
                    TabIndex =4
                    Name ="Text72"
                    ControlSource ="ClubAddress2"

                End
                Begin TextBox
                    Left =4920
                    Top =1440
                    Width =3420
                    Height =270
                    TabIndex =5
                    Name ="Text81"
                    ControlSource ="ClubCityST"

                End
                Begin TextBox
                    Left =5460
                    Top =1680
                    Width =1260
                    Height =300
                    TabIndex =6
                    Name ="Text74"
                    ControlSource ="ClubPhone1"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin Rectangle
                    Left =4800
                    Top =60
                    Width =3720
                    Height =2040
                    Name ="Box75"
                End
                Begin TextBox
                    Left =300
                    Top =420
                    Width =3420
                    Height =270
                    FontWeight =700
                    TabIndex =7
                    Name ="Text90"
                    ControlSource ="Name"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =180
                            Top =60
                            Width =3720
                            Height =300
                            BackColor =16053492
                            Name ="Label91"
                            Caption ="    Shipped from "
                        End
                    End
                End
                Begin TextBox
                    Left =300
                    Top =660
                    Width =3420
                    Height =270
                    TabIndex =8
                    Name ="Text92"
                    ControlSource ="=\"Agent \" & [AgentNumber]"

                End
                Begin TextBox
                    Left =300
                    Top =900
                    Width =3420
                    Height =270
                    TabIndex =9
                    Name ="Text93"
                    ControlSource ="Add1"

                End
                Begin TextBox
                    Left =300
                    Top =1200
                    Width =3420
                    Height =270
                    TabIndex =10
                    Name ="Text94"
                    ControlSource ="Add2"

                End
                Begin TextBox
                    Left =300
                    Top =1440
                    Width =3420
                    Height =270
                    TabIndex =11
                    Name ="Text95"
                    ControlSource ="Add3"

                End
                Begin Rectangle
                    Left =180
                    Top =60
                    Width =3720
                    Height =2040
                    Name ="Box97"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =5940
                    Top =2220
                    Width =720
                    Height =300
                    Name ="Label82"
                    Caption ="Line"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =900
                    Top =2220
                    Width =1140
                    Height =300
                    Name ="Label83"
                    Caption ="Item"
                    LayoutCachedLeft =900
                    LayoutCachedTop =2220
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =2520
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2040
                    Top =2220
                    Width =3060
                    Height =300
                    Name ="Label84"
                    Caption ="Description"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =165
                    Top =2220
                    Width =735
                    Height =300
                    Name ="Label85"
                    Caption ="Qty"
                    LayoutCachedLeft =165
                    LayoutCachedTop =2220
                    LayoutCachedWidth =900
                    LayoutCachedHeight =2520
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8460
                    Top =2220
                    Width =1140
                    Height =300
                    Name ="Label87"
                    Caption ="Price"
                    LayoutCachedLeft =8460
                    LayoutCachedTop =2220
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =2520
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =9600
                    Top =2220
                    Width =1020
                    Height =300
                    Name ="Label88"
                    Caption ="Ext Price"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =5100
                    Top =2220
                    Width =840
                    Height =300
                    Name ="Label177"
                    Caption ="Size"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =9600
                    Top =900
                    Width =720
                    TabIndex =12
                    Name ="txtPages"
                    ControlSource ="=[Pages]"

                End
                Begin TextBox
                    Left =840
                    Top =1680
                    Width =1260
                    Height =300
                    TabIndex =13
                    Name ="Text139"
                    ControlSource ="Phone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =2040
                    Top =1680
                    Width =420
                    Height =300
                    TabIndex =14
                    Name ="Text142"
                    ControlSource ="=\"Fax: \""

                End
                Begin TextBox
                    TextAlign =1
                    Left =2460
                    Top =1680
                    Width =1260
                    Height =300
                    TabIndex =15
                    Name ="Text143"
                    ControlSource ="Adv3"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =6660
                    Top =1680
                    Width =420
                    Height =300
                    TabIndex =16
                    Name ="Text144"
                    ControlSource ="=\"Fax: \""

                End
                Begin TextBox
                    TextAlign =1
                    Left =7080
                    Top =1680
                    Width =1260
                    Height =300
                    TabIndex =17
                    Name ="Text145"
                    ControlSource ="ClubFax"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =300
                    Top =1680
                    Width =480
                    Height =300
                    TabIndex =18
                    Name ="Text146"
                    ControlSource ="=\"Voice: \""

                End
                Begin TextBox
                    Left =4920
                    Top =1680
                    Width =480
                    Height =300
                    TabIndex =19
                    Name ="Text147"
                    ControlSource ="=\"Voice: \""

                End
                Begin Line
                    Left =180
                    Top =2520
                    Width =0
                    Height =60
                    Name ="Line167"
                End
                Begin Line
                    Left =900
                    Top =2520
                    Width =0
                    Height =60
                    Name ="Line168"
                    LayoutCachedLeft =900
                    LayoutCachedTop =2520
                    LayoutCachedWidth =900
                    LayoutCachedHeight =2580
                End
                Begin Line
                    Left =2040
                    Top =2520
                    Width =0
                    Height =60
                    Name ="Line169"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =5100
                    Top =2520
                    Width =0
                    Height =60
                    Name ="Line170"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =5940
                    Top =2520
                    Width =0
                    Height =60
                    Name ="Line171"
                End
                Begin Line
                    Left =6660
                    Top =2520
                    Width =0
                    Height =60
                    Name ="Line172"
                End
                Begin Line
                    Left =8760
                    Top =2520
                    Width =0
                    Height =60
                    Name ="Line173"
                End
                Begin Line
                    Left =9600
                    Top =2520
                    Width =0
                    Height =60
                    Name ="Line174"
                End
                Begin Line
                    Left =10620
                    Top =2520
                    Width =0
                    Height =60
                    Name ="Line175"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =8760
                    Top =1620
                    TabIndex =20
                    Name ="OuncesColumn"
                    ControlSource ="OuncesColumn"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7500
                    Top =2220
                    Width =960
                    Height =300
                    FontSize =11
                    FontWeight =700
                    TabIndex =21
                    ForeColor =8388608
                    Name ="Text180"
                    ControlSource ="=\"Net $/\" & [OuncesColumn] & \" oz\""
                    FontName ="Times New Roman"

                    LayoutCachedLeft =7500
                    LayoutCachedTop =2220
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =2520
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =6660
                    Top =2220
                    Width =840
                    Height =300
                    Name ="Label181"
                    Caption ="Net"
                End
                Begin Line
                    Left =7500
                    Top =2520
                    Width =0
                    Height =60
                    Name ="Line182"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    BackStyle =1
                    IMESentenceMode =3
                    Left =240
                    Width =10320
                    Height =255
                    TabIndex =9
                    Name ="Text181"

                    LayoutCachedLeft =240
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =8580
                    Width =735
                    Height =255
                    TabIndex =4
                    Name ="Price"
                    ControlSource ="NetPrice"
                    Format ="Standard"

                    LayoutCachedLeft =8580
                    LayoutCachedWidth =9315
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    DecimalPlaces =2
                    BackStyle =1
                    Left =9360
                    Width =180
                    TabIndex =7
                    Name ="SalePrice"
                    ControlSource ="SalePrice"
                    Format ="Standard"

                    LayoutCachedLeft =9360
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    BackStyle =1
                    Left =6000
                    Width =600
                    Height =255
                    Name ="LineID"
                    ControlSource ="LineID"

                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    Left =900
                    Width =1080
                    Height =225
                    TabIndex =1
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =900
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    BackStyle =1
                    Left =2100
                    Width =2880
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =240
                    Width =600
                    Height =255
                    FontWeight =700
                    TabIndex =3
                    Name ="QuantityOrdered"
                    ControlSource ="QuantityShipped"

                    LayoutCachedLeft =240
                    LayoutCachedWidth =840
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =9660
                    Width =900
                    Height =255
                    TabIndex =5
                    Name ="txtExtendCost"
                    ControlSource ="XPrice"
                    Format ="Standard"

                End
                Begin Line
                    LineSlant = NotDefault
                    Left =5100
                    Width =0
                    Height =255
                    Name ="Line91"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =840
                    Width =0
                    Height =255
                    Name ="Line92"
                    LayoutCachedLeft =840
                    LayoutCachedWidth =840
                    LayoutCachedHeight =255
                End
                Begin Line
                    Left =2040
                    Width =0
                    Height =255
                    Name ="Line93"
                End
                Begin Line
                    Left =6660
                    Width =0
                    Height =255
                    Name ="Line96"
                End
                Begin Line
                    Left =10620
                    Width =0
                    Height =255
                    Name ="Line98"
                End
                Begin Line
                    Left =180
                    Width =0
                    Height =255
                    Name ="Line133"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5160
                    Width =720
                    Height =255
                    TabIndex =6
                    Name ="SizeDesc"
                    ControlSource ="SizeDesc"

                End
                Begin Line
                    LineSlant = NotDefault
                    Left =5940
                    Width =0
                    Height =255
                    Name ="Line178"
                End
                Begin Line
                    Left =9600
                    Width =0
                    Height =255
                    Name ="Line185"
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =7560
                    Width =840
                    TabIndex =8
                    Name ="Text131"
                    ControlSource ="PricePerOZ"

                    LayoutCachedLeft =7560
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =240
                End
                Begin Line
                    Left =8460
                    Width =0
                    Height =255
                    Name ="Line132"
                    LayoutCachedLeft =8460
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    DecimalPlaces =3
                    TextAlign =3
                    BackStyle =1
                    Left =6720
                    Width =795
                    TabIndex =10
                    Name ="Text183"
                    ControlSource ="DiscountedPrice"
                    Format ="Standard"

                    LayoutCachedLeft =6720
                    LayoutCachedWidth =7515
                    LayoutCachedHeight =240
                End
                Begin Line
                    Left =7500
                    Width =0
                    Height =255
                    Name ="Line184"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =420
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =5835
                    Top =60
                    Width =1305
                    Height =300
                    FontSize =10
                    Name ="Text148"
                    ControlSource ="=Sum([XPrice])"
                    Format ="Standard"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =5100
                            Top =60
                            Width =645
                            Height =285
                            FontSize =10
                            ForeColor =0
                            Name ="Label149"
                            Caption ="Retail:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =7740
                    Top =60
                    Width =855
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="txtLiquorDiscount"
                    ControlSource ="LiquorDiscount"
                    Format ="Standard"

                    LayoutCachedLeft =7740
                    LayoutCachedTop =60
                    LayoutCachedWidth =8595
                    LayoutCachedHeight =345
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =7200
                            Top =60
                            Width =570
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label151"
                            Caption ="Disct:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =9540
                    Top =60
                    Width =1035
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="txtLiquorNetPrice"
                    ControlSource ="LiquorNetPrice"
                    Format ="Standard"

                    LayoutCachedLeft =9540
                    LayoutCachedTop =60
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =8655
                            Top =60
                            Width =915
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label153"
                            Caption ="Net Price:"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =4395
                    Top =60
                    Width =645
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="Text154"
                    ControlSource ="=Sum([QuantityShipped])"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =3420
                            Top =60
                            Width =855
                            Height =285
                            FontSize =10
                            ForeColor =0
                            Name ="Label155"
                            Caption ="Quantity"
                        End
                    End
                End
                Begin Rectangle
                    Left =180
                    Width =10440
                    Height =420
                    Name ="Box156"
                End
                Begin Label
                    TextAlign =3
                    Left =1500
                    Top =60
                    Width =1815
                    Height =300
                    Name ="lblLiquorOrRelated"
                    Caption ="Liquor Subtotal"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =7860
                    Top =60
                    Width =750
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="txtNonLiquorDiscount"
                    ControlSource ="NonLiquorDiscount"
                    Format ="Standard"

                    LayoutCachedLeft =7860
                    LayoutCachedTop =60
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =345
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =7200
                            Top =60
                            Width =570
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label161"
                            Caption ="Disct:"
                            LayoutCachedLeft =7200
                            LayoutCachedTop =60
                            LayoutCachedWidth =7770
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =9645
                    Top =60
                    Width =900
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="txtNonLiquorNetPrice"
                    ControlSource ="NonLiquorNetPrice"
                    Format ="Standard"

                    LayoutCachedLeft =9645
                    LayoutCachedTop =60
                    LayoutCachedWidth =10545
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =8640
                            Top =60
                            Width =915
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label163"
                            Caption ="Net Price:"
                            LayoutCachedLeft =8640
                            LayoutCachedTop =60
                            LayoutCachedWidth =9555
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin CheckBox
                    Left =1620
                    Top =60
                    TabIndex =6
                    Name ="chkLiquorItem"
                    ControlSource ="LiquorItem"

                End
            End
        End
        Begin PageFooter
            Height =585
            OnFormat ="[Event Procedure]"
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Top =300
                    Width =1920
                    Height =255
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text48"
                    ControlSource ="=Now()"
                    Format ="Long Time"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5700
                    Top =300
                    Width =4905
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Width =4740
                    Height =255
                    FontSize =9
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
                    TextAlign =3
                    TextFontFamily =18
                    Left =5700
                    Width =4905
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text64"
                    ControlSource ="=[Name]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Top =360
                    Width =10620
                    Height =180
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
                    TextAlign =2
                    TextFontFamily =18
                    Left =120
                    Width =10545
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"

                End
                Begin Label
                    TextAlign =2
                    Left =1440
                    Top =360
                    Width =3045
                    Height =180
                    FontSize =6
                    Name ="Label199"
                    Caption =" We can supply you with a full purchase history by month... "
                End
                Begin Label
                    TextAlign =2
                    Left =6360
                    Top =360
                    Width =2670
                    Height =180
                    FontSize =6
                    Name ="Label200"
                    Caption =" Request report #24 Club History by Bottle by Month"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =480
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =6075
                    Top =120
                    Width =1065
                    Height =300
                    FontSize =10
                    Name ="Text122"
                    ControlSource ="ExtPrice"
                    Format ="Standard"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =5340
                            Top =120
                            Width =645
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label122"
                            Caption ="Retail:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =7740
                    Top =120
                    Width =855
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Text98"
                    ControlSource ="Discount"
                    Format ="Standard"

                    LayoutCachedLeft =7740
                    LayoutCachedTop =120
                    LayoutCachedWidth =8595
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =7200
                            Top =120
                            Width =570
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label124"
                            Caption ="Disct:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =9540
                    Top =120
                    Width =1035
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Text99"
                    ControlSource ="ExtNetPrice"
                    Format ="Standard"

                    LayoutCachedLeft =9540
                    LayoutCachedTop =120
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =8640
                            Top =120
                            Width =930
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label39"
                            Caption ="Net Price:"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    Left =4740
                    Top =120
                    Width =540
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="Text102"
                    ControlSource ="=Sum([QuantityShipped])"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =3945
                            Top =120
                            Width =795
                            Height =300
                            FontSize =10
                            ForeColor =0
                            Name ="Label100"
                            Caption ="Quantity"
                        End
                    End
                End
                Begin Rectangle
                    Left =180
                    Top =60
                    Width =10440
                    Height =420
                    Name ="Box126"
                End
                Begin Label
                    Left =180
                    Top =120
                    Width =2880
                    Height =300
                    Name ="Label130"
                    Caption ="* Items priced '*' are on sale."
                End
                Begin Label
                    TextAlign =3
                    Left =3180
                    Top =120
                    Width =615
                    Height =300
                    Name ="Label166"
                    Caption ="Total:"
                End
            End
        End
    End
End
CodeBehindForm
' See "ClubInvoiceLiquorAndNonLiquor.cls"
