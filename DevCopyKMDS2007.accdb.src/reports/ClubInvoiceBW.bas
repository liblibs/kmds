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
    Width =10740
    DatasheetFontHeight =10
    ItemSuffix =139
    OnNoData ="[Event Procedure]"
    OnPage ="[Event Procedure]"
    OrderBy ="LineID"
    RecSrcDt = Begin
        0x63c794844aa9e240
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
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =4200
            OnFormat ="[Event Procedure]"
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =7680
                    Top =60
                    Width =3045
                    Height =510
                    FontSize =20
                    ForeColor =0
                    Name ="Label50"
                    Caption ="Club Invoice"
                    FontName ="Trebuchet MS"
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =7560
                    Height =510
                    FontSize =20
                    FontWeight =700
                    Name ="Text56"
                    ControlSource ="ClubName"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =600
                    Width =7560
                    Height =225
                    TabIndex =1
                    Name ="Text66"
                    ControlSource ="ClubContact"

                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =840
                    Width =7560
                    Height =225
                    TabIndex =2
                    Name ="Text67"
                    ControlSource ="ClubAddress1"

                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1080
                    Width =7560
                    Height =225
                    TabIndex =3
                    Name ="Text68"
                    ControlSource ="ClubAddress2"

                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1320
                    Width =7560
                    Height =225
                    TabIndex =4
                    Name ="Text69"
                    ControlSource ="ClubCityST"

                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =9060
                    Top =2580
                    Width =1545
                    Height =300
                    TabIndex =5
                    Name ="Text105"
                    ControlSource ="[ExtPrice]"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =9060
                    Top =2280
                    Width =1545
                    Height =300
                    BackColor =16053492
                    ForeColor =0
                    Name ="Label137"
                    Caption ="Retail"
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =7680
                    Top =600
                    Width =3060
                    Height =480
                    FontSize =20
                    FontWeight =700
                    TabIndex =6
                    Name ="Text70"
                    ControlSource ="InvoiceNo"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =9060
                    Top =3240
                    Width =1545
                    Height =300
                    TabIndex =7
                    Name ="Text71"
                    ControlSource ="Discount"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =9060
                    Top =2940
                    Width =1545
                    Height =300
                    BackColor =16053492
                    ForeColor =0
                    Name ="Label72"
                    Caption ="Discount"
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =9060
                    Top =1440
                    Width =1545
                    Height =420
                    FontSize =12
                    FontWeight =700
                    TabIndex =8
                    Name ="Text73"
                    ControlSource ="ExtNetPrice"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =9060
                    Top =1140
                    Width =1545
                    Height =300
                    BackColor =16053492
                    ForeColor =0
                    Name ="Label74"
                    Caption ="Invoice Amount"
                End
                Begin CheckBox
                    Visible = NotDefault
                    Left =240
                    Top =3300
                    TabIndex =9
                    Name ="NonLiquor"
                    ControlSource ="NonLiquor"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =240
                    Top =3240
                    Width =4500
                    TabIndex =10
                    Name ="txtSignature"
                    ControlSource ="=\"Signature\""

                    Begin
                        Begin Label
                            TextAlign =0
                            Left =240
                            Top =1740
                            Width =4440
                            Height =1380
                            ForeColor =0
                            Name ="Label117"
                            Caption ="I certify that this purchase will be used exclusively at a Dispenser Class A, B,"
                                " or C licensed business.  I understand that misrepresentation is illegal and can"
                                " subject me to penalties."
                        End
                    End
                End
                Begin Line
                    Left =1140
                    Top =3480
                    Width =3360
                    Name ="linSignature"
                End
                Begin Subform
                    OldBorderStyle =1
                    Left =4800
                    Top =1680
                    Width =4141
                    Height =1920
                    TabIndex =11
                    Name ="qryClubOrderSizeTotals"
                    SourceObject ="Report.ClubOrderSizeTotals"

                End
                Begin Rectangle
                    Left =180
                    Top =1680
                    Width =4560
                    Height =1920
                    Name ="Box121"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =180
                    Top =3660
                    Width =10440
                    Height =420
                    TabIndex =12
                    Name ="txtOrderNotes"
                    ControlSource ="OrderNotes"

                End
            End
        End
        Begin PageHeader
            Height =3360
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
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =9060
                            Top =60
                            Width =1545
                            Height =300
                            BackColor =16053492
                            ForeColor =0
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
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =180
                            Top =60
                            Width =1395
                            Height =300
                            ForeColor =0
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
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =6900
                            Top =720
                            Width =3720
                            Height =300
                            BackColor =16053492
                            ForeColor =0
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
                    Left =7020
                    Top =2340
                    Width =3420
                    Height =270
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
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =180
                            Top =720
                            Width =3720
                            Height =300
                            BackColor =16053492
                            ForeColor =0
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
                Begin TextBox
                    Left =300
                    Top =2340
                    Width =3420
                    Height =270
                    TabIndex =13
                    Name ="Text96"
                    ControlSource ="Phone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

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
                    ForeColor =0
                    Name ="Label82"
                    Caption ="Line"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1320
                    Top =3000
                    Width =720
                    Height =300
                    ForeColor =0
                    Name ="Label83"
                    Caption ="Item"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2040
                    Top =3000
                    Width =4200
                    Height =300
                    ForeColor =0
                    Name ="Label84"
                    Caption ="Description"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =180
                    Top =3000
                    Width =1140
                    Height =300
                    ForeColor =0
                    Name ="Label85"
                    Caption ="Quantity"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8760
                    Top =3000
                    Width =840
                    Height =300
                    ForeColor =0
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
                    ForeColor =0
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
                    ForeColor =0
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
                    ForeColor =0
                    Name ="Label133"
                    Caption ="Price/OZ"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5220
                    Top =240
                    Width =720
                    TabIndex =14
                    Name ="txtPages"
                    ControlSource ="=[Pages]"

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
                    BackStyle =1
                    Left =1380
                    Width =600
                    Height =255
                    TabIndex =1
                    Name ="ItemID"
                    ControlSource ="ItemID"

                End
                Begin TextBox
                    BackStyle =1
                    Left =2100
                    Width =4080
                    Height =255
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =240
                    Width =1020
                    Height =255
                    FontWeight =700
                    TabIndex =3
                    Name ="QuantityOrdered"
                    ControlSource ="QuantityShipped"

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
                    ControlSource ="Price"
                    Format ="Standard"

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
                    Left =1320
                    Width =0
                    Height =255
                    Name ="Line92"
                End
                Begin Line
                    Left =2040
                    Width =0
                    Height =255
                    Name ="Line93"
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

                End
                Begin Line
                    Left =8760
                    Width =0
                    Height =255
                    Name ="Line132"
                End
            End
        End
        Begin PageFooter
            Visible = NotDefault
            Height =480
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
            Height =480
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =6315
                    Top =120
                    Width =825
                    Height =285
                    FontSize =10
                    Name ="Text122"
                    ControlSource ="ExtPrice"
                    Format ="Standard"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =5580
                            Top =120
                            Width =645
                            Height =285
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
                    Left =7860
                    Top =120
                    Width =750
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text98"
                    ControlSource ="Discount"
                    Format ="Standard"

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
                    Left =9660
                    Top =120
                    Width =900
                    Height =255
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Text99"
                    ControlSource ="ExtNetPrice"
                    Format ="Standard"

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
                    Left =4875
                    Top =120
                    Width =645
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="Text102"
                    ControlSource ="=Sum([QuantityShipped])"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =3900
                            Top =120
                            Width =855
                            Height =285
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
                    Width =3210
                    Height =300
                    ForeColor =0
                    Name ="Label130"
                    Caption ="* Items priced '*' are on sale."
                End
            End
        End
    End
End
CodeBehindForm
' See "ClubInvoiceBW.cls"
