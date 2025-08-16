Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =9
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10680
    DatasheetFontHeight =10
    ItemSuffix =238
    Left =315
    Top =1200
    OnNoData ="[Event Procedure]"
    OnPage ="[Event Procedure]"
    RecSrcDt = Begin
        0x978984cb9495e240
    End
    RecordSource ="qryPurchaseOrderAndSysfile"
    Caption ="PurchaseOrder"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
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
            GroupFooter = NotDefault
            ControlSource ="ID"
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="Repack"
        End
        Begin BreakLevel
            ControlSource ="ItemID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =2160
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =7620
                    Width =2985
                    Height =540
                    FontSize =20
                    Name ="Label58"
                    Caption ="Purchase Order"
                    FontName ="Trebuchet MS"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Width =7320
                    Height =540
                    ColumnOrder =0
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text59"
                    ControlSource ="Name"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Top =540
                    Width =7320
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text61"
                    ControlSource ="Add1"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Top =780
                    Width =7320
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text64"
                    ControlSource ="Add2"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Top =1020
                    Width =7320
                    ColumnOrder =3
                    TabIndex =3
                    Name ="Text65"
                    ControlSource ="Add3"

                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =8460
                    Top =1620
                    Width =720
                    ColumnOrder =4
                    TabIndex =4
                    Name ="txtPages"
                    ControlSource ="=[Pages]"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7620
                    Top =540
                    Width =3000
                    Height =540
                    ColumnOrder =5
                    FontSize =20
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="Text210"
                    ControlSource ="ID"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =2640
                    Top =1740
                    Width =1395
                    Height =300
                    ColumnOrder =6
                    TabIndex =6
                    Name ="Text101"
                    ControlSource ="=Sum([cases])"

                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =9300
                    Top =1740
                    Width =1380
                    Height =300
                    ColumnOrder =7
                    TabIndex =7
                    Name ="Text105"
                    ControlSource ="=Sum([QuantityShipped]*[Price])"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =9300
                    Top =1440
                    Width =1380
                    Height =300
                    BackColor =16053492
                    Name ="Label137"
                    Caption ="Retail"
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =2640
                    Top =1440
                    Width =1395
                    Height =300
                    BackColor =16053492
                    Name ="lblQuantity"
                    Caption ="Cases"
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =240
                    Top =1740
                    Width =1380
                    Height =300
                    ColumnOrder =8
                    TabIndex =8
                    Name ="Text186"
                    ControlSource ="=Sum([QuantityShipped])"

                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =240
                    Top =1440
                    Width =1380
                    Height =300
                    BackColor =16053492
                    Name ="Label187"
                    Caption ="Quantity"
                End
                Begin CheckBox
                    Visible = NotDefault
                    SpecialEffect =2
                    Left =4620
                    Top =1560
                    Height =210
                    ColumnOrder =9
                    TabIndex =9
                    Name ="PoAtCost"
                    ControlSource ="qryPurchaseOrder.PoAtCost"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =0
                            TextFontFamily =34
                            Left =4860
                            Top =1500
                            Width =2415
                            Height =525
                            FontSize =8
                            FontWeight =400
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label53"
                            Caption ="DO NOT Print Alternating Gray Lines on Purchase orders"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
            End
        End
        Begin PageHeader
            Height =3540
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    Left =360
                    Top =1320
                    Width =3420
                    Height =270
                    ColumnWidth =2550
                    Name ="Name"
                    ControlSource ="VendorName"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =240
                            Top =960
                            Width =3720
                            Height =300
                            BackColor =16053492
                            Name ="Label67"
                            Caption ="    Vendor "
                        End
                    End
                End
                Begin TextBox
                    Left =360
                    Top =1560
                    Width =3420
                    Height =270
                    TabIndex =1
                    Name ="Contact"
                    ControlSource ="Contact"

                End
                Begin TextBox
                    Left =360
                    Top =1800
                    Width =3420
                    Height =270
                    TabIndex =2
                    Name ="Address1"
                    ControlSource ="Address1"

                End
                Begin TextBox
                    Left =360
                    Top =2100
                    Width =3420
                    Height =270
                    TabIndex =3
                    Name ="Address2"
                    ControlSource ="Address2"

                End
                Begin TextBox
                    Left =360
                    Top =2340
                    Width =3420
                    Height =270
                    TabIndex =4
                    Name ="CitySTZ"
                    ControlSource ="CitySTZ"

                End
                Begin TextBox
                    Left =360
                    Top =2580
                    Width =1680
                    Height =300
                    TabIndex =5
                    Name ="Phone1"
                    ControlSource ="Phone1"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin Rectangle
                    Left =240
                    Top =960
                    Width =3720
                    Height =2040
                    Name ="Box57"
                End
                Begin TextBox
                    Left =7080
                    Top =1320
                    Width =3420
                    Height =270
                    TabIndex =6
                    Name ="Text68"
                    ControlSource ="Name"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =6960
                            Top =960
                            Width =3720
                            Height =300
                            BackColor =16053492
                            Name ="Label69"
                            Caption ="    Ship to"
                        End
                    End
                End
                Begin TextBox
                    Left =7080
                    Top =1560
                    Width =3420
                    Height =270
                    TabIndex =7
                    Name ="Text70"
                    ControlSource ="Contact"

                End
                Begin TextBox
                    Left =7080
                    Top =1800
                    Width =3420
                    Height =270
                    TabIndex =8
                    Name ="Text71"
                    ControlSource ="Add1"

                End
                Begin TextBox
                    Left =7080
                    Top =2100
                    Width =3420
                    Height =270
                    TabIndex =9
                    Name ="Text72"
                    ControlSource ="Add2"

                End
                Begin TextBox
                    Left =7080
                    Top =2340
                    Width =3420
                    Height =270
                    TabIndex =10
                    Name ="Text73"
                    ControlSource ="Add3"

                End
                Begin TextBox
                    Left =7080
                    Top =2580
                    Width =3420
                    Height =270
                    TabIndex =11
                    Name ="Text74"
                    ControlSource ="Phone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin Rectangle
                    Left =6960
                    Top =960
                    Width =3720
                    Height =2040
                    Name ="Box75"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =7140
                    Top =3180
                    Width =720
                    Height =300
                    Name ="Label82"
                    Caption ="Line"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1380
                    Top =3180
                    Width =720
                    Height =300
                    Name ="Label83"
                    Caption ="Item"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2100
                    Top =3180
                    Width =4200
                    Height =300
                    Name ="Label84"
                    Caption ="Description"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =240
                    Top =3180
                    Width =1140
                    Height =300
                    Name ="Label85"
                    Caption ="Quantity"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8820
                    Top =3180
                    Width =840
                    Height =300
                    Name ="Label87"
                    Caption ="Price"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =9660
                    Top =3180
                    Width =1020
                    Height =300
                    Name ="Label88"
                    Caption ="Ext Price"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =6300
                    Top =3180
                    Width =840
                    Height =300
                    Name ="Label177"
                    Caption ="Size"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =7860
                    Top =3180
                    Width =960
                    Height =300
                    Name ="Label183"
                    Caption ="Cases"
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =240
                    Top =420
                    Width =1380
                    Height =300
                    TabIndex =12
                    Name ="Text76"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =240
                            Top =120
                            Width =1380
                            Height =300
                            BackColor =16053492
                            Name ="Label77"
                            Caption ="P.O. No."
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2640
                    Top =420
                    Width =1380
                    Height =300
                    TabIndex =13
                    Name ="Text79"
                    ControlSource ="OrderDate"
                    Format ="Medium Date"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =2640
                            Top =120
                            Width =1380
                            Height =300
                            BackColor =16053492
                            Name ="Label80"
                            Caption ="Date"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9285
                    Top =420
                    Width =1395
                    Height =300
                    TabIndex =14
                    Name ="Text212"
                    ControlSource ="=[Page] & \" of \" & [Pages]"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =9285
                            Top =120
                            Width =1395
                            Height =300
                            BackColor =16053492
                            Name ="Label213"
                            Caption ="Page Number"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4200
                    Top =120
                    Width =4980
                    Height =600
                    FontSize =24
                    TabIndex =15
                    Name ="txtPosted"

                End
                Begin TextBox
                    Left =2580
                    Top =2580
                    Width =1200
                    Height =300
                    TabIndex =16
                    Name ="FAX"
                    ControlSource ="qryPurchaseOrder.FAX"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    Begin
                        Begin Label
                            Left =2040
                            Top =2520
                            Width =540
                            Height =240
                            Name ="Label252"
                            Caption ="Fax:"
                        End
                    End
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
            Height =240
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    BackStyle =1
                    IMESentenceMode =3
                    Left =240
                    Width =10440
                    TabIndex =8
                    Name ="Text181"

                End
                Begin TextBox
                    BackStyle =1
                    Left =7200
                    Width =600
                    Height =225
                    Name ="LineID"
                    ControlSource ="LineID"

                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =1440
                    Width =600
                    Height =225
                    TabIndex =1
                    Name ="ItemID"
                    ControlSource ="ItemID"

                End
                Begin TextBox
                    BackStyle =1
                    Left =2160
                    Width =4080
                    Height =225
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =360
                    Width =825
                    FontWeight =700
                    TabIndex =3
                    Name ="QuantityOrdered"
                    ControlSource ="QuantityShipped"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =8880
                    Width =720
                    TabIndex =4
                    Name ="Price"
                    ControlSource ="Price"
                    Format ="Standard"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =9720
                    Width =900
                    TabIndex =5
                    Name ="ExtPrice"
                    ControlSource ="=[QuantityShipped]*[Price]"
                    Format ="Standard"

                End
                Begin Line
                    Left =6300
                    Width =0
                    Height =240
                    Name ="Line91"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1380
                    Width =0
                    Height =240
                    Name ="Line92"
                End
                Begin Line
                    Left =2100
                    Width =0
                    Height =240
                    Name ="Line93"
                End
                Begin Line
                    Left =8820
                    Width =0
                    Height =240
                    Name ="Line96"
                End
                Begin Line
                    Left =9660
                    Width =0
                    Height =240
                    Name ="Line98"
                End
                Begin Line
                    Left =10680
                    Width =0
                    Height =240
                    Name ="Line99"
                End
                Begin Line
                    Left =7860
                    Width =0
                    Height =240
                    Name ="Line133"
                End
                Begin CheckBox
                    Visible = NotDefault
                    TabIndex =6
                    Name ="Repack"
                    ControlSource ="Repack"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6360
                    Width =720
                    Height =225
                    TabIndex =7
                    Name ="SizeDesc"
                    ControlSource ="SizeDesc"

                End
                Begin Line
                    Left =7140
                    Width =0
                    Height =240
                    Name ="Line178"
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =7920
                    Width =825
                    TabIndex =9
                    Name ="Text184"
                    ControlSource ="Cases"
                    Format ="0;–0;\"\""

                End
                Begin Line
                    Left =240
                    Width =0
                    Height =240
                    Name ="Line185"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanShrink = NotDefault
            Height =360
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =360
                    Top =60
                    Width =960
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Text231"
                    ControlSource ="=Sum([QuantityShipped])"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =9300
                    Top =60
                    Width =1320
                    Height =255
                    TabIndex =2
                    Name ="Text232"
                    ControlSource ="=Sum([QuantityShipped]*[Price])"
                    Format ="Standard"

                End
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextAlign =3
                    BackStyle =1
                    Left =7920
                    Top =60
                    Width =840
                    Height =255
                    TabIndex =3
                    Name ="Text233"
                    ControlSource ="=Sum([Cases])"
                    Format ="Standard"

                End
                Begin TextBox
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =240
                    Width =10440
                    Height =360
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="txtRepacks"
                    ControlSource ="=\"Repacks are listed above\""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =240
                    Width =10440
                    Height =360
                    FontSize =11
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="txtFullCase"
                    ControlSource ="=\"Full Cases are listed above\""
                    FontName ="Times New Roman"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =375
            Name ="GroupFooter1"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    Left =3660
                    Top =60
                    Width =2460
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label39"
                    Caption ="Purchase Order Total"
                End
                Begin Line
                    Left =240
                    Top =360
                    Width =10440
                    Name ="Line117"
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =360
                    Top =60
                    Width =960
                    FontSize =10
                    FontWeight =700
                    Name ="Text119"
                    ControlSource ="=Sum([QuantityShipped])"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =9300
                    Top =60
                    Width =1320
                    Height =255
                    TabIndex =1
                    Name ="Text122"
                    ControlSource ="=Sum([QuantityShipped]*[Price])"
                    Format ="Standard"

                End
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextAlign =3
                    BackStyle =1
                    Left =7920
                    Top =60
                    Width =840
                    Height =255
                    TabIndex =2
                    Name ="Text123"
                    ControlSource ="=Sum([Cases])"
                    Format ="Standard"

                End
                Begin Line
                    Left =240
                    Width =10440
                    Name ="Line198"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1380
                    Width =0
                    Height =360
                    Name ="Line223"
                End
                Begin Line
                    Left =2100
                    Width =0
                    Height =360
                    Name ="Line224"
                End
                Begin Line
                    Left =8820
                    Width =0
                    Height =360
                    Name ="Line225"
                End
                Begin Line
                    Left =10680
                    Width =0
                    Height =360
                    Name ="Line227"
                End
                Begin Line
                    Left =7860
                    Width =0
                    Height =360
                    Name ="Line228"
                End
                Begin Line
                    Left =7140
                    Width =0
                    Height =360
                    Name ="Line229"
                End
                Begin Line
                    Left =240
                    Width =0
                    Height =360
                    Name ="Line230"
                End
            End
        End
        Begin PageFooter
            Visible = NotDefault
            Height =780
            Name ="PageFooterSection"
            Begin
                Begin Line
                    Visible = NotDefault
                    Left =240
                    Top =120
                    Width =10440
                    Name ="Line200"
                End
                Begin Line
                    Visible = NotDefault
                    Left =5160
                    Width =0
                    Height =120
                    Name ="Line201"
                End
                Begin Line
                    Visible = NotDefault
                    LineSlant = NotDefault
                    Left =240
                    Width =0
                    Height =120
                    Name ="Line202"
                End
                Begin Line
                    Visible = NotDefault
                    Left =960
                    Width =0
                    Height =120
                    Name ="Line203"
                End
                Begin Line
                    Visible = NotDefault
                    Left =8820
                    Width =0
                    Height =120
                    Name ="Line204"
                End
                Begin Line
                    Visible = NotDefault
                    Left =9660
                    Width =0
                    Height =120
                    Name ="Line205"
                End
                Begin Line
                    Visible = NotDefault
                    Left =10680
                    Width =0
                    Height =120
                    Name ="Line206"
                End
                Begin Line
                    Visible = NotDefault
                    Left =6720
                    Width =0
                    Height =120
                    Name ="Line207"
                End
                Begin Line
                    Visible = NotDefault
                    Left =6000
                    Width =0
                    Height =120
                    Name ="Line208"
                End
                Begin Line
                    Visible = NotDefault
                    Left =7680
                    Width =0
                    Height =120
                    Name ="Line209"
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Top =180
                    Width =4320
                    Name ="txtDate"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7200
                    Top =420
                    Width =3480
                    TabIndex =1
                    Name ="txtCashier"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Top =420
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
                    Top =180
                    Width =3480
                    TabIndex =3
                    Name ="Text218"
                    ControlSource ="=[Name]"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Top =480
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
                    Top =180
                    Width =10680
                    TabIndex =5
                    Name ="txtStoreName"
                    Format ="Long Date"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "Copy of PurchaseOrderLiquor.cls"
