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
    Width =10680
    DatasheetFontHeight =10
    ItemSuffix =253
    Left =585
    Top =1320
    OnNoData ="[Event Procedure]"
    OnPage ="[Event Procedure]"
    OrderBy ="LineID"
    RecSrcDt = Begin
        0x00537ce8cd08e340
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
            Height =1560
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
                    Left =7980
                    Top =1140
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
            End
        End
        Begin PageHeader
            Height =3480
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
                    Width =1620
                    Height =270
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
                    Top =1680
                    Width =3420
                    Height =270
                    TabIndex =7
                    Name ="Text71"
                    ControlSource ="Add1"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =1680
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =1950
                End
                Begin TextBox
                    Left =7080
                    Top =1980
                    Width =3420
                    Height =270
                    TabIndex =8
                    Name ="Text72"
                    ControlSource ="Add2"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =1980
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =2250
                End
                Begin TextBox
                    Left =7080
                    Top =2220
                    Width =3420
                    Height =270
                    TabIndex =9
                    Name ="Text73"
                    ControlSource ="Add3"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =2220
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =2490
                End
                Begin TextBox
                    Left =7080
                    Top =2580
                    Width =3420
                    Height =270
                    TabIndex =10
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
                    Left =9000
                    Top =3180
                    Width =1680
                    Height =300
                    Name ="Label83"
                    Caption ="Item"
                    LayoutCachedLeft =9000
                    LayoutCachedTop =3180
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =3480
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4080
                    Top =3180
                    Width =4920
                    Height =300
                    Name ="Label84"
                    Caption ="Description"
                    LayoutCachedLeft =4080
                    LayoutCachedTop =3180
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =3480
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =285
                    Top =3180
                    Width =1755
                    Height =300
                    Name ="Label85"
                    Caption ="Qty Ordered"
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
                            Top =120
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
                    Left =2040
                    Top =3180
                    Width =2040
                    Height =300
                    Name ="Label245"
                    Caption =" Vendor Stk #"
                    LayoutCachedLeft =2040
                    LayoutCachedTop =3180
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =3480
                End
                Begin TextBox
                    Left =2520
                    Top =2580
                    Width =1260
                    Height =270
                    TabIndex =14
                    Name ="FAX"
                    ControlSource ="qryPurchaseOrder.FAX"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    Begin
                        Begin Label
                            Left =1980
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
        Begin Section
            KeepTogether = NotDefault
            Height =435
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    BackStyle =1
                    Left =4200
                    Top =60
                    Width =4800
                    Height =315
                    FontSize =12
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =60
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    BackStyle =1
                    Left =240
                    Top =60
                    Width =1020
                    Height =315
                    FontSize =12
                    TabIndex =2
                    Name ="QuantityOrdered"
                    ControlSource ="QuantityShipped"

                End
                Begin TextBox
                    BackStyle =1
                    IMESentenceMode =3
                    Left =2220
                    Top =60
                    Width =1860
                    Height =315
                    FontSize =12
                    TabIndex =3
                    Name ="StockNumber"
                    ControlSource ="StockNumber"

                End
                Begin Line
                    Left =240
                    Top =420
                    Width =10380
                    Name ="Line250"
                End
                Begin TextBox
                    OverlapFlags =4
                    Left =8940
                    Top =60
                    Width =1680
                    Height =315
                    FontSize =12
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8940
                    LayoutCachedTop =60
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =375
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
' See "PurchaseOrderForFax.cls"
