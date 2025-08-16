Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =9
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =258
    Left =615
    OnNoData ="[Event Procedure]"
    OnPage ="[Event Procedure]"
    OrderBy ="LineID"
    RecSrcDt = Begin
        0xad084c845544e340
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =2100
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
                    Left =5100
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
                    Left =9300
                    Top =1740
                    Width =1380
                    Height =300
                    ColumnOrder =6
                    TabIndex =6
                    Name ="txtExtPriceHeader"
                    ControlSource ="=Sum([ExtPrice])"
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
                    Name ="lblExtPriceHeader"
                    Caption ="Retail"
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =240
                    Top =1740
                    Width =1380
                    Height =300
                    ColumnOrder =7
                    TabIndex =7
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
                    Name ="lblQuantity"
                    Caption ="Quantity"
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =7800
                    Top =1740
                    Width =1380
                    Height =300
                    ColumnOrder =8
                    TabIndex =8
                    Name ="txtExtCostHeader"
                    ControlSource ="=Sum([ExtCost])"
                    Format ="$#,##0.00;($#,##0.00)"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =7800
                            Top =1440
                            Width =1380
                            Height =300
                            BackColor =16053492
                            Name ="Label249"
                            Caption ="Cost"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =6360
                    Top =1740
                    Width =1380
                    Height =300
                    ColumnOrder =9
                    TabIndex =9
                    Name ="txtExtGPHeader"
                    ControlSource ="=(Sum([ExtPrice])-Sum([ExtCost]))/Sum([ExtPrice])"
                    Format ="Percent"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =6345
                            Top =1440
                            Width =1395
                            Height =300
                            BackColor =16053492
                            Name ="Label251"
                            Caption ="GP"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    SpecialEffect =2
                    Left =1860
                    Top =1560
                    Height =210
                    ColumnOrder =10
                    TabIndex =10
                    Name ="PoAtCost"
                    ControlSource ="qryPurchaseOrder.PoAtCost"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =0
                            TextFontFamily =34
                            Left =2100
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
            Height =3840
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    Left =375
                    Top =1140
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
                            Left =255
                            Top =780
                            Width =3720
                            Height =300
                            BackColor =16053492
                            Name ="Label67"
                            Caption ="    Vendor "
                        End
                    End
                End
                Begin TextBox
                    Left =375
                    Top =1380
                    Width =3420
                    Height =270
                    TabIndex =1
                    Name ="Contact"
                    ControlSource ="Contact"

                End
                Begin TextBox
                    Left =375
                    Top =1620
                    Width =3420
                    Height =270
                    TabIndex =2
                    Name ="Address1"
                    ControlSource ="Address1"

                End
                Begin TextBox
                    Left =375
                    Top =1920
                    Width =3420
                    Height =270
                    TabIndex =3
                    Name ="Address2"
                    ControlSource ="Address2"

                End
                Begin TextBox
                    Left =375
                    Top =2160
                    Width =3420
                    Height =270
                    TabIndex =4
                    Name ="CitySTZ"
                    ControlSource ="CitySTZ"

                End
                Begin TextBox
                    Left =375
                    Top =2400
                    Width =1620
                    Height =300
                    TabIndex =5
                    Name ="Phone1"
                    ControlSource ="Phone1"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin Rectangle
                    Left =255
                    Top =780
                    Width =3720
                    Height =2040
                    Name ="Box57"
                End
                Begin TextBox
                    Left =7095
                    Top =1140
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
                            Left =6975
                            Top =780
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
                    Top =1440
                    Width =3420
                    Height =270
                    TabIndex =7
                    Name ="Text71"
                    ControlSource ="Add1"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =1440
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =1710
                End
                Begin TextBox
                    Left =7080
                    Top =1740
                    Width =3420
                    Height =270
                    TabIndex =8
                    Name ="Text72"
                    ControlSource ="Add2"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =1740
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =2010
                End
                Begin TextBox
                    Left =7080
                    Top =1980
                    Width =3420
                    Height =270
                    TabIndex =9
                    Name ="Text73"
                    ControlSource ="Add3"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =1980
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =2250
                End
                Begin TextBox
                    Left =7095
                    Top =2400
                    Width =1560
                    Height =300
                    TabIndex =10
                    Name ="Text74"
                    ControlSource ="Phone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin Rectangle
                    Left =6975
                    Top =780
                    Width =3720
                    Height =2040
                    Name ="Box75"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8400
                    Top =3540
                    Width =540
                    Height =300
                    Name ="GPLabel"
                    Caption ="GP%"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1680
                    Top =3540
                    Width =1200
                    Height =300
                    Name ="Label83"
                    Caption ="Item"
                    LayoutCachedLeft =1680
                    LayoutCachedTop =3540
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =3840
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2880
                    Top =3540
                    Width =2220
                    Height =300
                    Name ="Label84"
                    Caption ="Description"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =5100
                    Top =3540
                    Width =600
                    Height =300
                    Name ="Label85"
                    Caption ="Qty"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =240
                    Top =3540
                    Width =840
                    Height =300
                    Name ="lblPrice"
                    Caption ="Cases"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1080
                    Top =3540
                    Width =600
                    Height =300
                    FontSize =9
                    Name ="lblExtPrice"
                    Caption ="Repax"
                    LayoutCachedLeft =1080
                    LayoutCachedTop =3540
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =3840
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =240
                    Top =360
                    Width =1380
                    Height =300
                    TabIndex =11
                    Name ="Text76"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =240
                            Top =60
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
                    Left =2580
                    Top =360
                    Width =1380
                    Height =300
                    TabIndex =12
                    Name ="Text79"
                    ControlSource ="OrderDate"
                    Format ="Medium Date"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =2580
                            Top =60
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
                    Top =360
                    Width =1395
                    Height =300
                    TabIndex =13
                    Name ="Text212"
                    ControlSource ="=[Page] & \" of \" & [Pages]"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =9285
                            Top =60
                            Width =1395
                            Height =300
                            BackColor =16053492
                            Name ="Label213"
                            Caption ="Page Number"
                        End
                    End
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8940
                    Top =3540
                    Width =855
                    Height =300
                    Name ="Label235"
                    Caption ="Cost"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =9780
                    Top =3540
                    Width =915
                    Height =300
                    Name ="Label236"
                    Caption ="Ext Cost"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =5700
                    Top =3540
                    Width =1320
                    Height =300
                    Name ="Label245"
                    Caption ="Vendor Stk #"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =7020
                    Top =3540
                    Width =1380
                    Height =300
                    Name ="Label246"
                    Caption ="UPC"
                End
                Begin TextBox
                    Left =9195
                    Top =2400
                    Height =270
                    TabIndex =14
                    Name ="Adv3"
                    ControlSource ="Adv3"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            TextFontFamily =34
                            Left =8715
                            Top =2400
                            Width =480
                            Height =270
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label253"
                            Caption ="Fax"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4140
                    Top =180
                    Width =5100
                    Height =480
                    FontSize =20
                    FontWeight =600
                    TabIndex =15
                    Name ="txtPosted"

                End
                Begin TextBox
                    Left =2595
                    Top =2400
                    Width =1200
                    Height =300
                    TabIndex =16
                    Name ="FAX"
                    ControlSource ="qryPurchaseOrder.FAX"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    Begin
                        Begin Label
                            Left =2055
                            Top =2340
                            Width =540
                            Height =240
                            Name ="Label252"
                            Caption ="Fax:"
                        End
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =240
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    BackStyle =1
                    Left =8460
                    Width =420
                    Height =225
                    Name ="GPercent"
                    ControlSource ="=IIf([Price]<>0,([Price]-[Cost])/[Price],0)"
                    Format ="Percent"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    BackStyle =1
                    Left =2940
                    Width =2100
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =5160
                    Width =480
                    FontWeight =700
                    TabIndex =3
                    Name ="QuantityOrdered"
                    ControlSource ="QuantityShipped"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    BackStyle =1
                    Left =300
                    Width =660
                    TabIndex =4
                    Name ="Cases"
                    ControlSource ="Cases"
                    Format ="#,##0;-#,##0[Red];\"\""

                    LayoutCachedLeft =300
                    LayoutCachedWidth =960
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    BackStyle =1
                    Left =1140
                    Width =540
                    TabIndex =5
                    Name ="Singles"
                    ControlSource ="qryPurchaseOrder.Singles"
                    Format ="#,##0;-#,##0[Red];\"\""

                    LayoutCachedLeft =1140
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =240
                End
                Begin Line
                    Left =8400
                    Width =0
                    Height =240
                    Name ="Line91"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1800
                    Width =0
                    Height =240
                    Name ="Line92"
                    LayoutCachedLeft =1800
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =240
                End
                Begin Line
                    Left =2880
                    Width =0
                    Height =240
                    Name ="Line93"
                End
                Begin Line
                    Left =5100
                    Width =0
                    Height =240
                    Name ="Line96"
                End
                Begin Line
                    Left =1080
                    Width =0
                    Height =240
                    Name ="linPrice"
                End
                Begin Line
                    Left =10680
                    Width =0
                    Height =240
                    Name ="linExtPrice"
                End
                Begin Line
                    Left =240
                    Width =0
                    Height =240
                    Name ="Line185"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =9000
                    Width =720
                    TabIndex =6
                    Name ="Cost"
                    ControlSource ="Cost"
                    Format ="Standard"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =9840
                    Width =780
                    TabIndex =7
                    Name ="ExtCost"
                    ControlSource ="ExtCost"
                    Format ="Standard"

                End
                Begin Line
                    Left =8940
                    Width =0
                    Height =240
                    Name ="Line239"
                End
                Begin Line
                    Left =9780
                    Width =0
                    Height =240
                    Name ="Line240"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5760
                    Width =1200
                    TabIndex =8
                    Name ="StockNumber"
                    ControlSource ="StockNumber"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7080
                    Width =1260
                    TabIndex =9
                    Name ="Text243"
                    ControlSource ="UPC"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    LineSlant = NotDefault
                    Left =5700
                    Width =0
                    Height =240
                    Name ="Line244"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =7020
                    Width =0
                    Height =240
                    Name ="Line247"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =10680
                    Width =0
                    Height =225
                    TabIndex =10
                    Name ="ExtPrice"
                    ControlSource ="ExtPrice"

                    LayoutCachedLeft =10680
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =2
                    Left =1680
                    Width =1320
                    TabIndex =1
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1680
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =240
                End
            End
        End
        Begin PageFooter
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
            Height =135
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin Line
                    Visible = NotDefault
                    Left =360
                    Top =120
                    Width =10440
                    Name ="Line200"
                End
                Begin Line
                    Visible = NotDefault
                    Left =1920
                    Width =0
                    Height =120
                    Name ="Line201"
                    LayoutCachedLeft =1920
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =120
                End
                Begin Line
                    Visible = NotDefault
                    LineSlant = NotDefault
                    Left =360
                    Width =0
                    Height =120
                    Name ="Line202"
                End
                Begin Line
                    Visible = NotDefault
                    Left =1200
                    Width =0
                    Height =120
                    Name ="Line203"
                End
                Begin Line
                    Visible = NotDefault
                    Left =8520
                    Width =0
                    Height =120
                    Name ="Line204"
                End
                Begin Line
                    Visible = NotDefault
                    Left =9060
                    Width =0
                    Height =120
                    Name ="Line205"
                End
                Begin Line
                    Visible = NotDefault
                    Left =10800
                    Width =0
                    Height =120
                    Name ="Line206"
                End
                Begin Line
                    Visible = NotDefault
                    Left =5220
                    Width =0
                    Height =120
                    Name ="Line207"
                End
                Begin Line
                    Visible = NotDefault
                    Left =3000
                    Width =0
                    Height =120
                    Name ="Line208"
                End
                Begin Line
                    Visible = NotDefault
                    Left =7140
                    Width =0
                    Height =120
                    Name ="Line209"
                End
                Begin Line
                    Visible = NotDefault
                    Left =9900
                    Width =0
                    Height =120
                    Name ="Line256"
                End
            End
        End
    End
End
CodeBehindForm
' See "PurchaseOrderWithCaseQty.cls"
