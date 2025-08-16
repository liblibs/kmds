Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10665
    DatasheetFontHeight =10
    ItemSuffix =177
    Left =1710
    Top =45
    OnNoData ="[Event Procedure]"
    OnPage ="[Event Procedure]"
    RecSrcDt = Begin
        0x8ec4b7d0a7aee540
    End
    RecordSource ="qryClubInvoiceLiquorAndNonLiquorClubScreen"
    Caption ="ClubInvoiceLiquorAndNonLiquor ClubScreen"
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
                    ControlSource ="=\"Invoice #\" & [InvoiceNumber]"
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
                    ControlSource ="=Sum([xPrice])"
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
                    Top =1080
                    Width =4680
                    Height =600
                    BackColor =16053492
                    Name ="Label176"
                    Caption ="Combined Invoice Reprinted from History file"
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
                    TabIndex =6
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
                    TabIndex =7
                    Name ="Text92"
                    ControlSource ="=\"Agent \" & [AgentNumber]"

                End
                Begin TextBox
                    Left =300
                    Top =900
                    Width =3420
                    Height =270
                    TabIndex =8
                    Name ="Text93"
                    ControlSource ="Add1"

                End
                Begin TextBox
                    Left =300
                    Top =1200
                    Width =3420
                    Height =270
                    TabIndex =9
                    Name ="Text94"
                    ControlSource ="Add2"

                End
                Begin TextBox
                    Left =300
                    Top =1440
                    Width =3420
                    Height =270
                    TabIndex =10
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
                    Left =7080
                    Top =2220
                    Width =720
                    Height =300
                    Name ="Label82"
                    Caption ="Line"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1020
                    Top =2220
                    Width =1380
                    Height =300
                    Name ="Label83"
                    Caption ="Item"
                    LayoutCachedLeft =1020
                    LayoutCachedTop =2220
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =2520
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2400
                    Top =2220
                    Width =3840
                    Height =300
                    Name ="Label84"
                    Caption ="Description"
                    LayoutCachedLeft =2400
                    LayoutCachedTop =2220
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =2520
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =165
                    Top =2220
                    Width =855
                    Height =300
                    Name ="Label85"
                    Caption ="Qty"
                    LayoutCachedLeft =165
                    LayoutCachedTop =2220
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =2520
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8760
                    Top =2220
                    Width =840
                    Height =300
                    Name ="Label87"
                    Caption ="Price"
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
                    Left =6240
                    Top =2220
                    Width =840
                    Height =300
                    Name ="Label177"
                    Caption ="Size"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =7800
                    Top =2220
                    Width =960
                    Height =300
                    Name ="Label133"
                    Caption ="Price/OZ"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =9600
                    Top =900
                    Width =720
                    TabIndex =11
                    Name ="txtPages"
                    ControlSource ="=[Pages]"

                End
                Begin TextBox
                    Left =840
                    Top =1680
                    Width =1260
                    Height =300
                    TabIndex =12
                    Name ="Text139"
                    ControlSource ="Phone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =2040
                    Top =1680
                    Width =420
                    Height =300
                    TabIndex =13
                    Name ="Text142"
                    ControlSource ="=\"Fax: \""

                End
                Begin TextBox
                    TextAlign =1
                    Left =2460
                    Top =1680
                    Width =1260
                    Height =300
                    TabIndex =14
                    Name ="Text143"
                    ControlSource ="Adv3"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =6660
                    Top =1680
                    Width =420
                    Height =300
                    TabIndex =15
                    Name ="Text144"
                    ControlSource ="=\"Fax: \""

                End
                Begin TextBox
                    TextAlign =1
                    Left =7080
                    Top =1680
                    Width =1260
                    Height =300
                    TabIndex =16
                    Name ="Text145"
                    ControlSource ="ClubFax"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =300
                    Top =1680
                    Width =480
                    Height =300
                    TabIndex =17
                    Name ="Text146"
                    ControlSource ="=\"Voice: \""

                End
                Begin TextBox
                    Left =4920
                    Top =1680
                    Width =480
                    Height =300
                    TabIndex =18
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
                    Left =1020
                    Top =2520
                    Width =0
                    Height =60
                    Name ="Line168"
                    LayoutCachedLeft =1020
                    LayoutCachedTop =2520
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =2580
                End
                Begin Line
                    Left =2400
                    Top =2520
                    Width =0
                    Height =60
                    Name ="Line169"
                    LayoutCachedLeft =2400
                    LayoutCachedTop =2520
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =2580
                End
                Begin Line
                    Left =6240
                    Top =2520
                    Width =0
                    Height =60
                    Name ="Line170"
                End
                Begin Line
                    Left =7080
                    Top =2520
                    Width =0
                    Height =60
                    Name ="Line171"
                End
                Begin Line
                    Left =7800
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
            End
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
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    BackStyle =1
                    Left =7140
                    Width =600
                    Height =255
                    Name ="LineID"
                    ControlSource ="LineID"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =1
                    BackStyle =1
                    Left =1080
                    Width =1260
                    Height =255
                    TabIndex =1
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1080
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    BackStyle =1
                    Left =2460
                    Width =3720
                    Height =255
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2460
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =240
                    Width =720
                    Height =255
                    FontWeight =700
                    TabIndex =3
                    Name ="QuantityOrdered"
                    ControlSource ="QuantityShipped"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =240
                    LayoutCachedWidth =960
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =8820
                    Width =540
                    Height =255
                    TabIndex =4
                    Name ="Price"
                    ControlSource ="UNetPrice"
                    Format ="Standard"
                    FontName ="Arial Narrow"

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
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Left =6240
                    Width =0
                    Height =255
                    Name ="Line91"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1020
                    Width =0
                    Height =255
                    Name ="Line92"
                    LayoutCachedLeft =1020
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =255
                End
                Begin Line
                    Left =2400
                    Width =0
                    Height =255
                    Name ="Line93"
                    LayoutCachedLeft =2400
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =255
                End
                Begin Line
                    Left =7800
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
                    IMESentenceMode =3
                    Left =6300
                    Width =720
                    Height =255
                    TabIndex =6
                    Name ="SizeDesc"
                    ControlSource ="SizeDesc"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Left =7080
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
                    BackStyle =1
                    Left =8040
                    Width =600
                    Height =255
                    TabIndex =8
                    Name ="Text131"
                    ControlSource ="PricePerOZ"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Left =8760
                    Width =0
                    Height =255
                    Name ="Line132"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =420
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =6315
                    Top =60
                    Width =825
                    Height =285
                    FontSize =10
                    Name ="Text148"
                    ControlSource ="=Sum([XPrice])"
                    Format ="Standard"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =5580
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
                    Left =7860
                    Top =60
                    Width =750
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="txtLiquorDiscount"
                    ControlSource ="LiquorDiscount"
                    Format ="Standard"

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
                    Left =9660
                    Top =60
                    Width =900
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="txtLiquorNetPrice"
                    ControlSource ="LiquorNetPrice"
                    Format ="Standard"

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
                    Left =4875
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
                            Left =3900
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
                    Left =1980
                    Top =60
                    Width =1815
                    Height =300
                    Name ="lblLiquorOrRelated"
                    Caption ="SubTotal:"
                End
                Begin TextBox
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
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =9660
                    Top =60
                    Width =900
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="txtNonLiquorNetPrice"
                    ControlSource ="NonLiquorNetPrice"
                    Format ="Standard"

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
                            Name ="Label163"
                            Caption ="Net Price:"
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
                    Top =300
                    Width =10620
                    Height =255
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
                    Left =9735
                    Top =120
                    Width =825
                    Height =300
                    FontSize =10
                    Name ="Text122"
                    ControlSource ="=Sum([xPrice])"
                    Format ="Standard"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =9000
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
                    TextAlign =2
                    BackStyle =1
                    Left =4875
                    Top =120
                    Width =645
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Text102"
                    ControlSource ="=Sum([QuantityShipped])"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =3930
                            Top =120
                            Width =825
                            Height =270
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
                    Left =240
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
' See "ClubInvoiceLiquorAndNonLiquorForClubScreen.cls"
