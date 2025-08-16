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
    Width =10755
    DatasheetFontHeight =10
    ItemSuffix =179
    Left =585
    Top =1320
    OnPage ="[Event Procedure]"
    OrderBy ="LineID"
    RecSrcDt = Begin
        0xf308bd209995e240
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =2400
            OnFormat ="[Event Procedure]"
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =7620
                    Top =120
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
                    Left =120
                    Top =120
                    Width =7320
                    Height =540
                    ColumnOrder =0
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text59"
                    ControlSource ="Name"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =1440
                            Top =1680
                            Width =1380
                            Height =300
                            Name ="Label136"
                            Caption ="Quantity"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =120
                    Top =720
                    Width =7320
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text61"
                    ControlSource ="Add1"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =120
                    Top =960
                    Width =7320
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text64"
                    ControlSource ="Add2"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =120
                    Top =1200
                    Width =7320
                    ColumnOrder =3
                    TabIndex =3
                    Name ="Text65"
                    ControlSource ="Add3"

                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7200
                    Top =1980
                    Width =1380
                    Height =300
                    ColumnOrder =4
                    TabIndex =4
                    Name ="Text76"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =7200
                            Top =1680
                            Width =1380
                            Height =300
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
                    Left =5760
                    Top =1980
                    Width =1380
                    Height =300
                    ColumnOrder =5
                    TabIndex =5
                    Name ="Text79"
                    ControlSource ="OrderDate"
                    Format ="Medium Date"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =5760
                            Top =1680
                            Width =1380
                            Height =300
                            Name ="Label80"
                            Caption ="Date"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =1440
                    Top =1980
                    Width =1380
                    Height =300
                    ColumnOrder =6
                    TabIndex =6
                    Name ="Text101"
                    ControlSource ="=Sum([QuantityShipped])"

                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =4320
                    Top =1980
                    Width =1380
                    Height =300
                    ColumnOrder =7
                    TabIndex =7
                    Name ="Text103"
                    ControlSource ="=Sum([QuantityShipped]*[Cost])"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =2880
                    Top =1980
                    Width =1380
                    Height =300
                    ColumnOrder =8
                    TabIndex =8
                    Name ="Text105"
                    ControlSource ="=Sum([QuantityShipped]*[Price])"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =2880
                    Top =1680
                    Width =1380
                    Height =300
                    Name ="Label137"
                    Caption ="Retail"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =4320
                    Top =1680
                    Width =1380
                    Height =300
                    Name ="Label138"
                    Caption ="Cost"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =9000
                    Top =1680
                    Width =660
                    ColumnOrder =9
                    TabIndex =9
                    Name ="txtPages"
                    ControlSource ="=[Pages]"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =660
                    Width =2940
                    Height =480
                    ColumnOrder =10
                    FontSize =20
                    FontWeight =700
                    TabIndex =10
                    ForeColor =8388608
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5640
                    Top =1140
                    Width =4980
                    Height =480
                    ColumnOrder =11
                    FontSize =24
                    TabIndex =11
                    Name ="txtPosted"

                End
            End
        End
        Begin PageHeader
            Height =2700
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    Left =1140
                    Top =480
                    Width =3420
                    Height =270
                    ColumnWidth =2550
                    Name ="Name"
                    ControlSource ="VendorName"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            Left =1020
                            Top =120
                            Width =3720
                            Height =300
                            Name ="Label67"
                            Caption ="    Vendor "
                        End
                    End
                End
                Begin TextBox
                    Left =1140
                    Top =720
                    Width =3420
                    Height =270
                    TabIndex =1
                    Name ="Contact"
                    ControlSource ="Contact"

                End
                Begin TextBox
                    Left =1140
                    Top =960
                    Width =3420
                    Height =270
                    TabIndex =2
                    Name ="Address1"
                    ControlSource ="Address1"

                End
                Begin TextBox
                    Left =1140
                    Top =1260
                    Width =3420
                    Height =270
                    TabIndex =3
                    Name ="Address2"
                    ControlSource ="Address2"

                End
                Begin TextBox
                    Left =1140
                    Top =1500
                    Width =3420
                    Height =270
                    TabIndex =4
                    Name ="CitySTZ"
                    ControlSource ="CitySTZ"

                End
                Begin TextBox
                    Left =1140
                    Top =1740
                    Width =1920
                    Height =300
                    TabIndex =5
                    Name ="Phone1"
                    ControlSource ="Phone1"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin Rectangle
                    Left =1020
                    Top =120
                    Width =3720
                    Height =2040
                    Name ="Box57"
                End
                Begin TextBox
                    Left =6180
                    Top =480
                    Width =3420
                    Height =270
                    TabIndex =6
                    Name ="Text68"
                    ControlSource ="Name"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            Left =6060
                            Top =120
                            Width =3720
                            Height =300
                            Name ="Label69"
                            Caption ="    Ship to"
                        End
                    End
                End
                Begin TextBox
                    Left =6180
                    Top =780
                    Width =3420
                    Height =270
                    TabIndex =7
                    Name ="Text71"
                    ControlSource ="Add1"

                    LayoutCachedLeft =6180
                    LayoutCachedTop =780
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =1050
                End
                Begin TextBox
                    Left =6180
                    Top =1080
                    Width =3420
                    Height =270
                    TabIndex =8
                    Name ="Text72"
                    ControlSource ="Add2"

                    LayoutCachedLeft =6180
                    LayoutCachedTop =1080
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =1350
                End
                Begin TextBox
                    Left =6180
                    Top =1320
                    Width =3420
                    Height =270
                    TabIndex =9
                    Name ="Text73"
                    ControlSource ="Add3"

                    LayoutCachedLeft =6180
                    LayoutCachedTop =1320
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =1590
                End
                Begin TextBox
                    Left =6180
                    Top =1740
                    Width =3420
                    Height =270
                    TabIndex =10
                    Name ="Text74"
                    ControlSource ="Phone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin Rectangle
                    Left =6060
                    Top =120
                    Width =3720
                    Height =2040
                    Name ="Box75"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =360
                    Top =2400
                    Width =600
                    Height =300
                    Name ="Label82"
                    Caption ="Line"
                    LayoutCachedLeft =360
                    LayoutCachedTop =2400
                    LayoutCachedWidth =960
                    LayoutCachedHeight =2700
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =960
                    Top =2400
                    Width =1320
                    Height =300
                    Name ="Label83"
                    Caption ="Item"
                    LayoutCachedLeft =960
                    LayoutCachedTop =2400
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =2700
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2280
                    Top =2400
                    Width =3000
                    Height =300
                    Name ="Label84"
                    Caption ="Description"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =5280
                    Top =2400
                    Width =885
                    Height =300
                    Name ="Label85"
                    Caption ="Quantity"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8025
                    Top =2400
                    Width =855
                    Height =300
                    Name ="Label86"
                    Caption ="Cost"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =6180
                    Top =2400
                    Width =840
                    Height =300
                    Name ="Label87"
                    Caption ="Price"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =7020
                    Top =2400
                    Width =1020
                    Height =300
                    Name ="Label88"
                    Caption ="Ext Price"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8873
                    Top =2400
                    Width =960
                    Height =300
                    Name ="Label89"
                    Caption ="Ext Cost"
                End
                Begin TextBox
                    Left =3360
                    Top =1740
                    Width =1200
                    Height =300
                    TabIndex =11
                    Name ="FAX"
                    ControlSource ="qryPurchaseOrder.FAX"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    Begin
                        Begin Label
                            Left =2820
                            Top =1680
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
            Height =360
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =360
                    Width =600
                    Height =225
                    Name ="LineID"
                    ControlSource ="LineID"

                    LayoutCachedLeft =360
                    LayoutCachedWidth =960
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    BackStyle =1
                    Left =2340
                    Width =2880
                    Height =225
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="Description"

                    LayoutCachedLeft =2340
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =5280
                    Width =825
                    TabIndex =3
                    Name ="QuantityOrdered"
                    ControlSource ="QuantityShipped"

                    LayoutCachedLeft =5280
                    LayoutCachedWidth =6105
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =8100
                    Width =720
                    TabIndex =4
                    Name ="Cost"
                    ControlSource ="Cost"
                    Format ="Standard"

                    LayoutCachedLeft =8100
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =6240
                    Width =720
                    Height =255
                    TabIndex =5
                    Name ="Price"
                    ControlSource ="Price"
                    Format ="Standard"

                    LayoutCachedLeft =6240
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =7080
                    Width =900
                    Height =255
                    TabIndex =6
                    Name ="ExtPrice"
                    ControlSource ="ExtPrice"
                    Format ="Standard"

                    LayoutCachedLeft =7080
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    DecimalPlaces =2
                    BorderWidth =1
                    TextAlign =3
                    BackStyle =1
                    Left =8940
                    Width =840
                    Height =255
                    TabIndex =7
                    Name ="ExtCost"
                    ControlSource ="ExtCost"
                    Format ="Standard"

                    LayoutCachedLeft =8940
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =255
                End
                Begin Line
                    Left =960
                    Width =0
                    Height =360
                    Name ="Line91"
                    LayoutCachedLeft =960
                    LayoutCachedWidth =960
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =300
                    Width =0
                    Height =360
                    Name ="Line92"
                    LayoutCachedLeft =300
                    LayoutCachedWidth =300
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =2280
                    Width =0
                    Height =360
                    Name ="Line93"
                    LayoutCachedLeft =2280
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =6180
                    Width =0
                    Height =360
                    Name ="Line96"
                    LayoutCachedLeft =6180
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =7020
                    Width =0
                    Height =360
                    Name ="Line98"
                    LayoutCachedLeft =7020
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =8040
                    Width =0
                    Height =360
                    Name ="Line99"
                    LayoutCachedLeft =8040
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =8880
                    Width =0
                    Height =360
                    Name ="Line100"
                    LayoutCachedLeft =8880
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =5280
                    Width =0
                    Height =360
                    Name ="Line133"
                    LayoutCachedLeft =5280
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =9840
                    Width =0
                    Height =360
                    Name ="Line134"
                    LayoutCachedLeft =9840
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OverlapFlags =12
                    TextAlign =2
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
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =375
            Name ="GroupFooter1"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    Left =2340
                    Width =2460
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label39"
                    Caption ="Purchase Order Total"
                    LayoutCachedLeft =2340
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =270
                End
                Begin Line
                    Left =300
                    Top =360
                    Width =9540
                    Name ="Line117"
                    LayoutCachedLeft =300
                    LayoutCachedTop =360
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =5340
                    Top =60
                    Width =765
                    Name ="Text119"
                    ControlSource ="=Sum([QuantityShipped])"

                    LayoutCachedLeft =5340
                    LayoutCachedTop =60
                    LayoutCachedWidth =6105
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =7080
                    Top =60
                    Width =900
                    Height =255
                    TabIndex =1
                    Name ="Text122"
                    ControlSource ="=Sum([ExtPrice])"
                    Format ="Standard"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =60
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    DecimalPlaces =2
                    BorderWidth =1
                    TextAlign =3
                    BackStyle =1
                    Left =8400
                    Top =60
                    Width =1380
                    Height =255
                    TabIndex =2
                    Name ="Text123"
                    ControlSource ="=Sum([QuantityShipped]*[Cost])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8400
                    LayoutCachedTop =60
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =315
                End
                Begin Line
                    Left =300
                    Width =0
                    Height =360
                    Name ="Line124"
                    LayoutCachedLeft =300
                    LayoutCachedWidth =300
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =960
                    Width =0
                    Height =360
                    Name ="Line125"
                    LayoutCachedLeft =960
                    LayoutCachedWidth =960
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =2280
                    Width =0
                    Height =360
                    Name ="Line126"
                    LayoutCachedLeft =2280
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =5280
                    Width =0
                    Height =360
                    Name ="Line127"
                    LayoutCachedLeft =5280
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =6180
                    Width =0
                    Height =360
                    Name ="Line128"
                    LayoutCachedLeft =6180
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =7020
                    Width =0
                    Height =360
                    Name ="Line129"
                    LayoutCachedLeft =7020
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =9840
                    Width =0
                    Height =360
                    Name ="Line130"
                    LayoutCachedLeft =9840
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =8040
                    Width =0
                    Height =360
                    Name ="Line131"
                    LayoutCachedLeft =8040
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =8880
                    Width =0
                    Height =360
                    Name ="Line132"
                    LayoutCachedLeft =8880
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =720
                    Width =9120
                    Name ="Line135"
                    LayoutCachedLeft =720
                    LayoutCachedWidth =9840
                End
            End
        End
        Begin PageFooter
            Height =555
            Name ="PageFooterSection"
            Begin
                Begin Line
                    Left =300
                    Top =120
                    Width =9540
                    Name ="Line163"
                    LayoutCachedLeft =300
                    LayoutCachedTop =120
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =120
                End
                Begin Line
                    Left =300
                    Width =0
                    Height =120
                    Name ="Line164"
                    LayoutCachedLeft =300
                    LayoutCachedWidth =300
                    LayoutCachedHeight =120
                End
                Begin Line
                    Left =960
                    Width =0
                    Height =120
                    Name ="Line165"
                    LayoutCachedLeft =960
                    LayoutCachedWidth =960
                    LayoutCachedHeight =120
                End
                Begin Line
                    Left =2280
                    Width =0
                    Height =120
                    Name ="Line166"
                    LayoutCachedLeft =2280
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =120
                End
                Begin Line
                    Left =5280
                    Width =0
                    Height =120
                    Name ="Line167"
                    LayoutCachedLeft =5280
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =120
                End
                Begin Line
                    Left =6180
                    Width =0
                    Height =120
                    Name ="Line168"
                    LayoutCachedLeft =6180
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =120
                End
                Begin Line
                    Left =7020
                    Width =0
                    Height =120
                    Name ="Line169"
                    LayoutCachedLeft =7020
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =120
                End
                Begin Line
                    Left =9840
                    Width =0
                    Height =120
                    Name ="Line170"
                    LayoutCachedLeft =9840
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =120
                End
                Begin Line
                    Left =8040
                    Width =0
                    Height =120
                    Name ="Line171"
                    LayoutCachedLeft =8040
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =120
                End
                Begin Line
                    Left =8880
                    Width =0
                    Height =120
                    Name ="Line172"
                    LayoutCachedLeft =8880
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =120
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5640
                    Top =300
                    Width =5040
                    Height =255
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text177"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5640
                    LayoutCachedTop =300
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =555
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "PurchaseOrder.cls"
