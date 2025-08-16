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
    Width =10740
    DatasheetFontHeight =10
    ItemSuffix =196
    Left =315
    Top =1200
    OnPage ="[Event Procedure]"
    OrderBy ="LineID"
    RecSrcDt = Begin
        0x756ec6a11af0e240
    End
    RecordSource ="qryAgentTransferLinesAndSysfile"
    Caption ="AgentTransfer"
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
            Height =4140
            OnFormat ="[Event Procedure]"
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =7545
                    Top =120
                    Width =3060
                    Height =540
                    FontSize =20
                    Name ="lblTransfer"
                    Caption ="Transfer"
                    FontName ="Trebuchet MS"
                End
                Begin TextBox
                    TextAlign =1
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
                    ControlSource ="ThisAgentName"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =1440
                            Top =3540
                            Width =1380
                            Height =300
                            Name ="Label136"
                            Caption ="Quantity"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =120
                    Top =660
                    Width =7320
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text61"
                    ControlSource ="Add1"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =120
                    Top =900
                    Width =7320
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text64"
                    ControlSource ="Add2"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =120
                    Top =1140
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
                    Top =3840
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
                            Top =3540
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
                    Top =3840
                    Width =1380
                    Height =300
                    ColumnOrder =5
                    TabIndex =5
                    Name ="Text79"
                    ControlSource ="OrdDate"
                    Format ="Medium Date"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =5760
                            Top =3540
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
                    Top =3840
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
                    Left =2880
                    Top =3840
                    Width =1380
                    Height =300
                    ColumnOrder =7
                    TabIndex =7
                    Name ="Text105"
                    ControlSource ="=Sum([QuantityShipped]*[Price])"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =2880
                    Top =3540
                    Width =1380
                    Height =300
                    Name ="Label137"
                    Caption ="Retail"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =9000
                    Top =3840
                    Width =720
                    ColumnOrder =8
                    TabIndex =8
                    Name ="txtPages"
                    ControlSource ="=[Pages]"

                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =9000
                    Top =3540
                    Width =600
                    ColumnOrder =9
                    TabIndex =9
                    Name ="Status"
                    ControlSource ="Status"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7560
                    Top =660
                    Width =3060
                    Height =540
                    ColumnOrder =10
                    FontSize =20
                    FontWeight =700
                    TabIndex =10
                    ForeColor =8388608
                    Name ="Text175"
                    ControlSource ="=\"Agent #\" & [AgentNumber]"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Left =180
                    Top =1740
                    Width =3420
                    Height =270
                    ColumnWidth =2550
                    ColumnOrder =11
                    TabIndex =11
                    Name ="Name"
                    ControlSource ="=\"Agent #\" & [AgentID]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            Left =60
                            Top =1380
                            Width =3720
                            Height =300
                            Name ="lblShipper"
                            Caption ="   Ship from "
                        End
                    End
                End
                Begin TextBox
                    Left =180
                    Top =1980
                    Width =3420
                    Height =270
                    ColumnOrder =12
                    TabIndex =12
                    Name ="Contact"
                    ControlSource ="AgentName"

                End
                Begin TextBox
                    Left =180
                    Top =2220
                    Width =3420
                    Height =270
                    ColumnOrder =13
                    TabIndex =13
                    Name ="Address1"
                    ControlSource ="Address"

                End
                Begin TextBox
                    Left =180
                    Top =2520
                    Width =3420
                    Height =270
                    ColumnOrder =14
                    TabIndex =14
                    Name ="Address2"
                    ControlSource ="CitySTZ"

                End
                Begin TextBox
                    Left =180
                    Top =3000
                    Width =3420
                    Height =270
                    ColumnOrder =15
                    TabIndex =15
                    Name ="Phone1"
                    ControlSource ="qryAgentTransferLines.Phone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin Rectangle
                    Left =60
                    Top =1380
                    Width =3720
                    Height =2040
                    Name ="Box57"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4560
                    Top =1380
                    Width =6000
                    Height =600
                    ColumnOrder =16
                    FontSize =24
                    TabIndex =16
                    Name ="txtPosted"

                End
            End
        End
        Begin PageHeader
            Height =660
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =5700
                    Top =360
                    Width =660
                    Height =300
                    Name ="Label82"
                    Caption ="Line"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =900
                    Top =360
                    Width =1560
                    Height =300
                    Name ="Label83"
                    Caption ="Item"
                    LayoutCachedLeft =900
                    LayoutCachedTop =360
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =660
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2460
                    Top =360
                    Width =3240
                    Height =300
                    Name ="Label84"
                    Caption ="Description"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =300
                    Top =360
                    Width =600
                    Height =300
                    Name ="Label85"
                    Caption ="Qty"
                    LayoutCachedLeft =300
                    LayoutCachedTop =360
                    LayoutCachedWidth =900
                    LayoutCachedHeight =660
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =6330
                    Top =360
                    Width =885
                    Height =300
                    Name ="Label87"
                    Caption ="Cost"
                    LayoutCachedLeft =6330
                    LayoutCachedTop =360
                    LayoutCachedWidth =7215
                    LayoutCachedHeight =660
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =7200
                    Top =360
                    Width =1020
                    Height =300
                    Name ="Label88"
                    Caption ="Ext Cost"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8220
                    Top =360
                    Width =840
                    Height =300
                    Name ="Label184"
                    Caption ="Price"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =9060
                    Top =360
                    Width =1020
                    Height =300
                    Name ="Label185"
                    Caption ="Ext Price"
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
                    BackStyle =1
                    Left =2520
                    Width =2580
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="Description"

                    LayoutCachedLeft =2520
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =5880
                    Width =420
                    Name ="LineID"
                    ControlSource ="LineID"

                End
                Begin TextBox
                    BackStyle =1
                    Left =975
                    Width =1425
                    Height =225
                    TabIndex =1
                    Name ="ItemID"
                    ControlSource ="ItemID"

                    LayoutCachedLeft =975
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =300
                    Width =570
                    TabIndex =3
                    Name ="QuantityOrdered"
                    ControlSource ="QuantityShipped"

                    LayoutCachedLeft =300
                    LayoutCachedWidth =870
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =6420
                    Width =720
                    Height =255
                    TabIndex =4
                    Name ="Price"
                    ControlSource ="Cost"
                    Format ="Standard"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =7260
                    Width =900
                    Height =255
                    TabIndex =5
                    Name ="ExtPrice"
                    ControlSource ="ExtCost"
                    Format ="Standard"

                End
                Begin Line
                    Left =300
                    Width =0
                    Height =360
                    Name ="Line91"
                    LayoutCachedLeft =300
                    LayoutCachedWidth =300
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =900
                    Width =0
                    Height =360
                    Name ="Line92"
                    LayoutCachedLeft =900
                    LayoutCachedWidth =900
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =2460
                    Width =0
                    Height =360
                    Name ="Line93"
                End
                Begin Line
                    Left =6360
                    Width =0
                    Height =360
                    Name ="Line96"
                End
                Begin Line
                    Left =7200
                    Width =0
                    Height =360
                    Name ="Line98"
                End
                Begin Line
                    Left =8220
                    Width =0
                    Height =360
                    Name ="Line99"
                End
                Begin Line
                    Left =5700
                    Width =0
                    Height =360
                    Name ="Line133"
                End
                Begin TextBox
                    BackStyle =1
                    IMESentenceMode =3
                    Left =5100
                    Width =600
                    TabIndex =6
                    Name ="SizeDesc"
                    ControlSource ="SizeDesc"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =8280
                    Width =720
                    Height =255
                    TabIndex =7
                    Name ="Text186"
                    ControlSource ="Price"
                    Format ="Standard"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =9120
                    Width =900
                    Height =255
                    TabIndex =8
                    Name ="Text187"
                    ControlSource ="ExtPrice"
                    Format ="Standard"

                End
                Begin Line
                    Left =9060
                    Width =0
                    Height =360
                    Name ="Line188"
                End
                Begin Line
                    Left =10080
                    Width =0
                    Height =360
                    Name ="Line189"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =617
            Name ="GroupFooter1"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    Left =2520
                    Top =60
                    Width =2460
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label39"
                    Caption ="Transfer Total"
                End
                Begin Line
                    Left =300
                    Top =60
                    Width =9720
                    Name ="Line117"
                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =60
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =360
                    Top =60
                    Width =510
                    Name ="Text119"
                    ControlSource ="=Sum([QuantityShipped])"

                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =870
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =7260
                    Top =75
                    Width =900
                    Height =255
                    TabIndex =1
                    Name ="Text122"
                    ControlSource ="=Sum([ExtCost])"
                    Format ="Standard"

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
                    Left =900
                    Width =0
                    Height =360
                    Name ="Line125"
                    LayoutCachedLeft =900
                    LayoutCachedWidth =900
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =2460
                    Width =0
                    Height =360
                    Name ="Line126"
                End
                Begin Line
                    Left =5700
                    Width =0
                    Height =360
                    Name ="Line127"
                End
                Begin Line
                    Left =6360
                    Width =0
                    Height =360
                    Name ="Line128"
                End
                Begin Line
                    Left =7200
                    Width =0
                    Height =360
                    Name ="Line129"
                End
                Begin Line
                    Left =8220
                    Width =0
                    Height =360
                    Name ="Line131"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =9120
                    Top =75
                    Width =900
                    Height =255
                    TabIndex =2
                    Name ="Text190"
                    ControlSource ="=Sum([ExtPrice])"
                    Format ="Standard"

                End
                Begin Line
                    Left =9060
                    Width =0
                    Height =360
                    Name ="Line191"
                End
                Begin Line
                    Left =10080
                    Width =0
                    Height =360
                    Name ="Line192"
                    LayoutCachedLeft =10080
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =300
                    Top =360
                    Width =9780
                    Name ="Line193"
                    LayoutCachedLeft =300
                    LayoutCachedTop =360
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =360
                End
            End
        End
        Begin PageFooter
            Height =555
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
                    Width =5040
                    Height =255
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
                    Width =5040
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
                    Width =5040
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text194"
                    ControlSource ="=[Name]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Top =300
                    Width =10680
                    Height =255
                    FontSize =6
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="Text195"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =60
                    Width =10680
                    Height =255
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
            Height =795
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5580
                    Top =120
                    Width =1500
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="AgentNumber"
                    ControlSource ="AgentNumber"

                    Begin
                        Begin Label
                            TextAlign =0
                            Left =5580
                            Top =480
                            Width =1605
                            Height =300
                            Name ="Label178"
                            Caption ="Agent Signature"
                        End
                    End
                End
                Begin Line
                    Left =7260
                    Top =780
                    Width =3180
                    Name ="Line179"
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =120
                    Width =1500
                    Height =330
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="Text180"
                    ControlSource ="AgentID"

                    Begin
                        Begin Label
                            TextAlign =0
                            Left =60
                            Top =480
                            Width =1605
                            Height =300
                            Name ="Label181"
                            Caption ="Agent Signature"
                        End
                    End
                End
                Begin Line
                    Left =1740
                    Top =780
                    Width =3180
                    Name ="Line182"
                End
            End
        End
    End
End
CodeBehindForm
' See "AgentTransfer.cls"
