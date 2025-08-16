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
    ItemSuffix =176
    Left =345
    OnNoData ="[Event Procedure]"
    Filter ="(SalesDepartmentID=1)"
    OnPage ="[Event Procedure]"
    OrderBy ="LineID"
    RecSrcDt = Begin
        0xd8790758cbd3e240
    End
    RecordSource ="qryClubInvoiceAndSysfile"
    Caption ="Club Invoice"
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
            GroupFooter = NotDefault
            ControlSource ="LiquorItem"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =3060
            OnFormat ="[Event Procedure]"
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =1
                    BackStyle =1
                    Left =60
                    Top =480
                    Width =7320
                    Height =555
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
                    Left =7380
                    Top =480
                    Width =3240
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
                    Top =1380
                    Width =1620
                    Height =360
                    ColumnOrder =2
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="Text73"
                    ControlSource ="LiquorNetPrice"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =9000
                    Top =1080
                    Width =1620
                    Height =300
                    BackColor =16053492
                    Name ="Label74"
                    Caption ="Liquor Amount"
                End
                Begin CheckBox
                    Visible = NotDefault
                    Left =240
                    Top =2700
                    ColumnOrder =3
                    TabIndex =3
                    Name ="NonLiquor"
                    ControlSource ="NonLiquor"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =240
                    Top =2640
                    Width =4500
                    ColumnOrder =4
                    TabIndex =4
                    Name ="txtSignature"
                    ControlSource ="=\"Signature\""

                    Begin
                        Begin Label
                            TextAlign =0
                            Left =240
                            Top =1140
                            Width =4440
                            Height =1380
                            Name ="Label117"
                            Caption ="I certify that this purchase will be used exclusively at a Dispenser Class A, B,"
                                " or C licensed business.  I understand that misrepresentation is illegal and can"
                                " subject me to penalties."
                        End
                    End
                End
                Begin Line
                    Left =1140
                    Top =2880
                    Width =3360
                    Name ="linSignature"
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =1
                    Left =4800
                    Top =1080
                    Width =4141
                    Height =1920
                    TabIndex =5
                    Name ="qryClubOrderSizeTotals"
                    SourceObject ="Report.ClubOrderSizeTotals"

                End
                Begin Rectangle
                    Left =180
                    Top =1080
                    Width =4560
                    Height =1920
                    Name ="Box121"
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9000
                    Top =2760
                    Width =1620
                    ColumnOrder =5
                    TabIndex =6
                    Name ="Text140"
                    ControlSource ="InvDate"
                    Format ="Short Date"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =9000
                            Top =2460
                            Width =1620
                            Height =300
                            BackColor =16053492
                            Name ="Label141"
                            Caption ="Invoice Date"
                        End
                    End
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =2
                    Left =4560
                    Top =60
                    Width =1755
                    Height =420
                    FontSize =16
                    Name ="Label169"
                    Caption ="Store Copy"
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =9000
                    Top =2040
                    Width =1620
                    Height =360
                    ColumnOrder =6
                    FontSize =12
                    FontWeight =700
                    TabIndex =7
                    Name ="Text170"
                    ControlSource ="ExtNetPrice"
                    Format ="$#,##0.00;($#,##0.00)"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =9000
                            Top =1740
                            Width =1620
                            Height =300
                            BackColor =16053492
                            Name ="Label171"
                            Caption ="Invoice Amount"
                        End
                    End
                End
            End
        End
        Begin PageHeader
            Height =3360
            OnFormat ="[Event Procedure]"
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
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =180
                    Top =360
                    Width =1395
                    Height =300
                    TabIndex =1
                    Name ="Text75"
                    ControlSource ="InvDate"
                    Format ="Medium Date"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =180
                            Top =60
                            Width =1395
                            Height =300
                            BackColor =16053492
                            Name ="lblInvoiceDate"
                            Caption ="Invoice Date"
                        End
                    End
                End
                Begin TextBox
                    Left =7020
                    Top =1080
                    Width =3420
                    Height =270
                    TabIndex =2
                    Name ="Text77"
                    ControlSource ="ClubName"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =6900
                            Top =720
                            Width =3720
                            Height =300
                            BackColor =16053492
                            Name ="Label69"
                            Caption ="    Ship to / Sold to"
                        End
                    End
                End
                Begin TextBox
                    Left =7020
                    Top =1320
                    Width =3420
                    Height =270
                    TabIndex =3
                    Name ="Text78"
                    ControlSource ="ClubContact"

                End
                Begin TextBox
                    Left =7020
                    Top =1560
                    Width =3420
                    Height =270
                    TabIndex =4
                    Name ="Text80"
                    ControlSource ="ClubAddress1"

                End
                Begin TextBox
                    Left =7020
                    Top =1860
                    Width =3420
                    Height =270
                    TabIndex =5
                    Name ="Text72"
                    ControlSource ="ClubAddress2"

                End
                Begin TextBox
                    Left =7020
                    Top =2100
                    Width =3420
                    Height =270
                    TabIndex =6
                    Name ="Text81"
                    ControlSource ="ClubCityST"

                End
                Begin TextBox
                    Left =7560
                    Top =2340
                    Width =1260
                    Height =300
                    TabIndex =7
                    Name ="Text74"
                    ControlSource ="ClubPhone1"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin Rectangle
                    Left =6900
                    Top =720
                    Width =3720
                    Height =2040
                    Name ="Box75"
                End
                Begin TextBox
                    Left =300
                    Top =1080
                    Width =3420
                    Height =270
                    FontWeight =700
                    TabIndex =8
                    Name ="Text90"
                    ControlSource ="Name"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =180
                            Top =720
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
                    Top =1320
                    Width =3420
                    Height =270
                    TabIndex =9
                    Name ="Text92"
                    ControlSource ="=\"Agent \" & [AgentNumber]"

                End
                Begin TextBox
                    Left =300
                    Top =1560
                    Width =3420
                    Height =270
                    TabIndex =10
                    Name ="Text93"
                    ControlSource ="Add1"

                End
                Begin TextBox
                    Left =300
                    Top =1860
                    Width =3420
                    Height =270
                    TabIndex =11
                    Name ="Text94"
                    ControlSource ="Add2"

                End
                Begin TextBox
                    Left =300
                    Top =2100
                    Width =3420
                    Height =270
                    TabIndex =12
                    Name ="Text95"
                    ControlSource ="Add3"

                End
                Begin Rectangle
                    Left =180
                    Top =720
                    Width =3720
                    Height =2040
                    Name ="Box97"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =7080
                    Top =3000
                    Width =720
                    Height =300
                    Name ="Label82"
                    Caption ="Line"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =915
                    Top =3000
                    Width =1425
                    Height =300
                    Name ="Label83"
                    Caption ="Item"
                    LayoutCachedLeft =915
                    LayoutCachedTop =3000
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =3300
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2340
                    Top =3000
                    Width =3900
                    Height =300
                    Name ="Label84"
                    Caption ="Description"
                    LayoutCachedLeft =2340
                    LayoutCachedTop =3000
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =3300
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =165
                    Top =3000
                    Width =735
                    Height =300
                    Name ="Label85"
                    Caption ="Qty"
                    LayoutCachedLeft =165
                    LayoutCachedTop =3000
                    LayoutCachedWidth =900
                    LayoutCachedHeight =3300
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8760
                    Top =3000
                    Width =840
                    Height =300
                    Name ="Label87"
                    Caption ="Price"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =9600
                    Top =3000
                    Width =1020
                    Height =300
                    Name ="Label88"
                    Caption ="Ext Price"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =6240
                    Top =3000
                    Width =840
                    Height =300
                    Name ="Label177"
                    Caption ="Size"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =7800
                    Top =3000
                    Width =960
                    Height =300
                    Name ="Label133"
                    Caption ="Price/OZ"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5100
                    Top =900
                    Width =720
                    TabIndex =13
                    Name ="txtPages"
                    ControlSource ="=[Pages]"

                End
                Begin TextBox
                    Left =840
                    Top =2340
                    Width =1260
                    Height =300
                    TabIndex =14
                    Name ="Text139"
                    ControlSource ="Phone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =2040
                    Top =2340
                    Width =420
                    Height =300
                    TabIndex =15
                    Name ="Text142"
                    ControlSource ="=\"Fax: \""

                End
                Begin TextBox
                    TextAlign =1
                    Left =2460
                    Top =2340
                    Width =1260
                    Height =300
                    TabIndex =16
                    Name ="Text143"
                    ControlSource ="Adv3"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =8760
                    Top =2340
                    Width =420
                    Height =300
                    TabIndex =17
                    Name ="Text144"
                    ControlSource ="=\"Fax: \""

                End
                Begin TextBox
                    TextAlign =1
                    Left =9180
                    Top =2340
                    Width =1260
                    Height =300
                    TabIndex =18
                    Name ="Text145"
                    ControlSource ="ClubFax"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    Left =300
                    Top =2340
                    Width =480
                    Height =300
                    TabIndex =19
                    Name ="Text146"
                    ControlSource ="=\"Voice: \""

                End
                Begin TextBox
                    Left =7020
                    Top =2340
                    Width =480
                    Height =300
                    TabIndex =20
                    Name ="Text147"
                    ControlSource ="=\"Voice: \""

                End
                Begin TextBox
                    Visible = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    IMESentenceMode =3
                    Left =1680
                    Top =60
                    Width =7260
                    Height =300
                    ColumnWidth =1785
                    FontSize =11
                    FontWeight =700
                    TabIndex =21
                    BackColor =16053492
                    ForeColor =8388608
                    Name ="lblCombined"
                    ControlSource ="=\"See Customer Copy of Invoice for Total Amount Owed \" & Format([NonLiquorNetP"
                        "rice]+[LiquorNetPrice],\"Currency\")"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =4740
                    Top =420
                    ColumnWidth =1785
                    TabIndex =22
                    Name ="NonLiquorExtPrice"
                    ControlSource ="NonLiquorExtPrice"

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
                    Left =60
                    Width =10320
                    Height =255
                    TabIndex =9
                    Name ="Text181"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =10380
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

                End
                Begin TextBox
                    BackStyle =1
                    Left =7140
                    Width =600
                    Height =255
                    Name ="LineID"
                    ControlSource ="LineID"

                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =960
                    Width =1320
                    Height =225
                    TabIndex =1
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =960
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    BackStyle =1
                    Left =2400
                    Width =3780
                    Height =255
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="Description"

                    LayoutCachedLeft =2400
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =255
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
                    Left =8820
                    Width =735
                    Height =255
                    TabIndex =4
                    Name ="Price"
                    ControlSource ="NetPrice"
                    Format ="Standard"

                    LayoutCachedLeft =8820
                    LayoutCachedWidth =9555
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
                    Left =6240
                    Width =0
                    Height =255
                    Name ="Line91"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =900
                    Width =0
                    Height =255
                    Name ="Line92"
                    LayoutCachedLeft =900
                    LayoutCachedWidth =900
                    LayoutCachedHeight =255
                End
                Begin Line
                    Left =2340
                    Width =0
                    Height =255
                    Name ="Line93"
                    LayoutCachedLeft =2340
                    LayoutCachedWidth =2340
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

                End
                Begin Line
                    Left =7080
                    Width =0
                    Height =255
                    Name ="Line178"
                End
                Begin Line
                    Left =9660
                    Width =0
                    Height =255
                    Name ="Line185"
                    LayoutCachedLeft =9660
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    DecimalPlaces =2
                    BackStyle =1
                    Left =7980
                    Width =600
                    Height =255
                    TabIndex =8
                    Name ="Text131"
                    ControlSource ="PricePerOZ"
                    Format ="Fixed"

                    LayoutCachedLeft =7980
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =255
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
            Height =420
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =5535
                    Top =60
                    Width =1305
                    Height =300
                    FontSize =10
                    Name ="Text148"
                    ControlSource ="=Sum([XPrice])"
                    Format ="Standard"

                    LayoutCachedLeft =5535
                    LayoutCachedTop =60
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =4800
                            Top =60
                            Width =645
                            Height =285
                            FontSize =10
                            ForeColor =0
                            Name ="Label149"
                            Caption ="Retail:"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =60
                            LayoutCachedWidth =5445
                            LayoutCachedHeight =345
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =7560
                    Top =60
                    Width =915
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="txtLiquorDiscount"
                    ControlSource ="LiquorDiscount"
                    Format ="Standard"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =60
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =345
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =6900
                            Top =60
                            Width =570
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label151"
                            Caption ="Disct:"
                            LayoutCachedLeft =6900
                            LayoutCachedTop =60
                            LayoutCachedWidth =7470
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =9480
                    Top =60
                    Width =1095
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="txtLiquorNetPrice"
                    ControlSource ="LiquorNetPrice"
                    Format ="Standard"

                    LayoutCachedLeft =9480
                    LayoutCachedTop =60
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =8700
                            Top =60
                            Width =855
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label153"
                            Caption ="Net Price:"
                            LayoutCachedLeft =8700
                            LayoutCachedTop =60
                            LayoutCachedWidth =9555
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =4095
                    Top =60
                    Width =645
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="Text154"
                    ControlSource ="=Sum([QuantityShipped])"

                    LayoutCachedLeft =4095
                    LayoutCachedTop =60
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =345
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =3120
                            Top =60
                            Width =855
                            Height =285
                            FontSize =10
                            ForeColor =0
                            Name ="Label155"
                            Caption ="Quantity"
                            LayoutCachedLeft =3120
                            LayoutCachedTop =60
                            LayoutCachedWidth =3975
                            LayoutCachedHeight =345
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
                    Left =1200
                    Top =60
                    Width =1815
                    Height =300
                    Name ="lblLiquorOrRelated"
                    Caption ="Liquor Subtotal"
                    LayoutCachedLeft =1200
                    LayoutCachedTop =60
                    LayoutCachedWidth =3015
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =7560
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="txtNonLiquorDiscount"
                    ControlSource ="NonLiquorDiscount"
                    Format ="Standard"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =60
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =345
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =6900
                            Top =60
                            Width =570
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label161"
                            Caption ="Disct:"
                            LayoutCachedLeft =6900
                            LayoutCachedTop =60
                            LayoutCachedWidth =7470
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =9480
                    Top =60
                    Width =1095
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="txtNonLiquorNetPrice"
                    ControlSource ="NonLiquorNetPrice"
                    Format ="Standard"

                    LayoutCachedLeft =9480
                    LayoutCachedTop =60
                    LayoutCachedWidth =10575
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
                    Left =1320
                    Top =60
                    TabIndex =6
                    Name ="chkLiquorItem"
                    ControlSource ="LiquorItem"

                    LayoutCachedLeft =1320
                    LayoutCachedTop =60
                    LayoutCachedWidth =1580
                    LayoutCachedHeight =300
                End
            End
        End
        Begin PageFooter
            Height =0
            OnFormat ="[Event Procedure]"
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =1080
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =5535
                    Top =120
                    Width =1305
                    Height =300
                    FontSize =10
                    Name ="Text122"
                    ControlSource ="=Sum([xPrice])"
                    Format ="Standard"

                    LayoutCachedLeft =5535
                    LayoutCachedTop =120
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =4800
                            Top =120
                            Width =645
                            Height =285
                            FontSize =10
                            ForeColor =0
                            Name ="Label122"
                            Caption ="Retail:"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =120
                            LayoutCachedWidth =5445
                            LayoutCachedHeight =405
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =7560
                    Top =120
                    Width =960
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text98"
                    ControlSource ="Discount"
                    Format ="Standard"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =120
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =405
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =6900
                            Top =120
                            Width =570
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label124"
                            Caption ="Disct:"
                            LayoutCachedLeft =6900
                            LayoutCachedTop =120
                            LayoutCachedWidth =7470
                            LayoutCachedHeight =390
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
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Text99"
                    ControlSource ="ExtNetPrice"
                    Format ="Standard"

                    LayoutCachedLeft =9540
                    LayoutCachedTop =120
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =375
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =8655
                            Top =120
                            Width =915
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label39"
                            Caption ="Net Price:"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =4095
                    Top =120
                    Width =645
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="Text102"
                    ControlSource ="=Sum([QuantityShipped])"

                    LayoutCachedLeft =4095
                    LayoutCachedTop =120
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =405
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =3120
                            Top =120
                            Width =855
                            Height =285
                            FontSize =10
                            ForeColor =0
                            Name ="Label100"
                            Caption ="Quantity"
                            LayoutCachedLeft =3120
                            LayoutCachedTop =120
                            LayoutCachedWidth =3975
                            LayoutCachedHeight =405
                        End
                    End
                End
                Begin Rectangle
                    Left =180
                    Width =10440
                    Height =420
                    Name ="Box126"
                    LayoutCachedLeft =180
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =420
                End
                Begin Label
                    Left =240
                    Top =120
                    Width =2955
                    Height =300
                    Name ="Label130"
                    Caption ="* Items priced '*' are on sale."
                    LayoutCachedLeft =240
                    LayoutCachedTop =120
                    LayoutCachedWidth =3195
                    LayoutCachedHeight =420
                End
                Begin Label
                    OldBorderStyle =1
                    Left =180
                    Top =660
                    Width =2100
                    Height =420
                    Name ="Label172"
                    Caption ="Entered By: _______"
                End
                Begin Label
                    OldBorderStyle =1
                    Left =2400
                    Top =660
                    Width =2040
                    Height =420
                    Name ="Label173"
                    Caption ="Picked By: _______"
                End
                Begin Label
                    OldBorderStyle =1
                    Left =4560
                    Top =660
                    Width =2160
                    Height =420
                    Name ="Label174"
                    Caption ="Checked By: _______"
                End
                Begin Label
                    OldBorderStyle =1
                    Left =6840
                    Top =660
                    Width =3705
                    Height =420
                    Name ="Label175"
                    Caption ="Check # _____ Date: ____/____/____"
                End
            End
        End
    End
End
CodeBehindForm
' See "ClubInvoiceStoreCopy.cls"
