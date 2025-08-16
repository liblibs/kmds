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
    Width =10710
    DatasheetFontHeight =10
    ItemSuffix =143
    Left =1305
    Top =300
    OnNoData ="[Event Procedure]"
    OnPage ="[Event Procedure]"
    OrderBy ="LineID"
    RecSrcDt = Begin
        0x2b66fae9560be340
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
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1680
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
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =8520
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
                    TextAlign =2
                    Left =60
                    Top =600
                    Width =4440
                    Height =255
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text66"
                    ControlSource ="ClubContact"

                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =840
                    Width =4440
                    Height =255
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text67"
                    ControlSource ="ClubAddress1"

                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1080
                    Width =4440
                    Height =255
                    ColumnOrder =3
                    TabIndex =3
                    Name ="Text68"
                    ControlSource ="ClubAddress2"

                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1320
                    Width =4440
                    Height =255
                    ColumnOrder =4
                    TabIndex =4
                    Name ="Text69"
                    ControlSource ="ClubCityST"

                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =9840
                    Top =1200
                    Width =720
                    ColumnOrder =5
                    TabIndex =5
                    Name ="txtPages"
                    ControlSource ="=[Pages]"

                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =8715
                    Top =600
                    Width =1965
                    Height =540
                    ColumnOrder =6
                    FontSize =20
                    FontWeight =700
                    TabIndex =6
                    ForeColor =8388608
                    Name ="Text70"
                    ControlSource ="InvoiceNo"
                    FontName ="Trebuchet MS"

                End
                Begin CheckBox
                    Visible = NotDefault
                    Left =9120
                    Top =1260
                    ColumnOrder =7
                    TabIndex =7
                    Name ="NonLiquor"
                    ControlSource ="NonLiquor"

                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =1
                    Left =4560
                    Top =600
                    Width =4081
                    Height =1080
                    TabIndex =8
                    Name ="qryClubOrderSizeTotals"
                    SourceObject ="Report.ClubOrderSizeTotals"

                End
            End
        End
        Begin PageHeader
            Visible = NotDefault
            Height =2220
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9150
                    Top =1080
                    Width =1545
                    Height =300
                    Name ="Text76"
                    ControlSource ="InvoiceNo"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =9135
                            Top =780
                            Width =1575
                            Height =300
                            BackColor =13882323
                            Name ="Label77"
                            Caption ="Invoice No.."
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9150
                    Top =1800
                    Width =1545
                    Height =300
                    TabIndex =1
                    Name ="Text79"
                    ControlSource ="InvDate"
                    Format ="Medium Date"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =2
                            Left =9135
                            Top =1500
                            Width =1575
                            Height =300
                            BackColor =13882323
                            Name ="Label80"
                            Caption ="Invoice Date"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9120
                    Top =360
                    Width =1545
                    Height =300
                    TabIndex =2
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
                    Left =5295
                    Top =420
                    Width =3420
                    Height =270
                    TabIndex =3
                    Name ="Text77"
                    ControlSource ="ClubName"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =5175
                            Top =60
                            Width =3720
                            Height =300
                            BackColor =13882323
                            Name ="Label69"
                            Caption ="    Ship to / Sold to"
                        End
                    End
                End
                Begin TextBox
                    Left =5295
                    Top =660
                    Width =3420
                    Height =270
                    TabIndex =4
                    Name ="Text78"
                    ControlSource ="ClubContact"

                End
                Begin TextBox
                    Left =5295
                    Top =900
                    Width =3420
                    Height =270
                    TabIndex =5
                    Name ="Text80"
                    ControlSource ="ClubAddress1"

                End
                Begin TextBox
                    Left =5295
                    Top =1200
                    Width =3420
                    Height =270
                    TabIndex =6
                    Name ="Text72"
                    ControlSource ="ClubAddress2"

                End
                Begin TextBox
                    Left =5295
                    Top =1440
                    Width =3420
                    Height =270
                    TabIndex =7
                    Name ="Text81"
                    ControlSource ="ClubCityST"

                End
                Begin TextBox
                    Left =5295
                    Top =1680
                    Width =3420
                    Height =270
                    TabIndex =8
                    Name ="Text74"
                    ControlSource ="ClubPhone1"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin Rectangle
                    Left =5175
                    Top =60
                    Width =3720
                    Height =2040
                    Name ="Box75"
                End
                Begin TextBox
                    Left =315
                    Top =420
                    Width =3420
                    Height =270
                    FontWeight =700
                    TabIndex =9
                    Name ="Text90"
                    ControlSource ="Name"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =195
                            Top =60
                            Width =3720
                            Height =300
                            BackColor =13882323
                            Name ="Label91"
                            Caption ="    Shipped from "
                        End
                    End
                End
                Begin TextBox
                    Left =315
                    Top =660
                    Width =3420
                    Height =270
                    TabIndex =10
                    Name ="Text92"
                    ControlSource ="=\"Agent \" & [AgentNumber]"

                End
                Begin TextBox
                    Left =315
                    Top =900
                    Width =3420
                    Height =270
                    TabIndex =11
                    Name ="Text93"
                    ControlSource ="Add1"

                End
                Begin TextBox
                    Left =315
                    Top =1200
                    Width =3420
                    Height =270
                    TabIndex =12
                    Name ="Text94"
                    ControlSource ="Add2"

                End
                Begin TextBox
                    Left =315
                    Top =1440
                    Width =3420
                    Height =270
                    TabIndex =13
                    Name ="Text95"
                    ControlSource ="Add3"

                End
                Begin TextBox
                    Left =315
                    Top =1680
                    Width =3420
                    Height =270
                    TabIndex =14
                    Name ="Text96"
                    ControlSource ="Phone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin Rectangle
                    Left =195
                    Top =60
                    Width =3720
                    Height =2040
                    Name ="Box97"
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
                    OldBorderStyle =1
                    TextAlign =2
                    Left =600
                    Top =60
                    Width =1020
                    Height =300
                    Name ="Label83"
                    Caption ="Item"
                    LayoutCachedLeft =600
                    LayoutCachedTop =60
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =360
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1620
                    Top =60
                    Width =2640
                    Height =300
                    Name ="Label84"
                    Caption ="Description"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =45
                    Top =60
                    Width =555
                    Height =300
                    Name ="Label85"
                    Caption ="Qty"
                    LayoutCachedLeft =45
                    LayoutCachedTop =60
                    LayoutCachedWidth =600
                    LayoutCachedHeight =360
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4260
                    Top =60
                    Width =720
                    Height =300
                    Name ="Label177"
                    Caption ="Size"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =420
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    BackStyle =1
                    IMESentenceMode =3
                    Left =60
                    Width =4920
                    Height =420
                    FontSize =11
                    TabIndex =4
                    Name ="Text141"

                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    Left =480
                    Top =60
                    Width =1065
                    Height =285
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =480
                    LayoutCachedTop =60
                    LayoutCachedWidth =1545
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    BackStyle =1
                    Left =1680
                    Top =60
                    Width =2520
                    Height =345
                    FontSize =10
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =120
                    Top =60
                    Width =420
                    Height =360
                    TabIndex =2
                    Name ="QuantityOrdered"
                    ControlSource ="QuantityShipped"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =540
                    LayoutCachedHeight =420
                End
                Begin Line
                    Left =4980
                    Width =0
                    Height =420
                    Name ="Line91"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =60
                    Width =0
                    Height =420
                    Name ="Line92"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =600
                    Width =0
                    Height =420
                    Name ="Line93"
                    LayoutCachedLeft =600
                    LayoutCachedWidth =600
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    BackStyle =1
                    Left =4320
                    Top =60
                    Width =600
                    Height =345
                    FontSize =11
                    TabIndex =3
                    Name ="Text129"
                    ControlSource ="SizeDesc"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Left =1620
                    Width =0
                    Height =420
                    Name ="Line134"
                End
                Begin Line
                    Left =4260
                    Width =0
                    Height =420
                    Name ="Line135"
                End
            End
        End
        Begin PageFooter
            Visible = NotDefault
            Height =480
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

                    LayoutCachedWidth =4320
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7200
                    Top =240
                    Width =3480
                    TabIndex =1
                    Name ="txtCashier"

                    LayoutCachedLeft =7200
                    LayoutCachedTop =240
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =480
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

                    LayoutCachedTop =240
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7200
                    Width =3480
                    TabIndex =3
                    Name ="Text218"
                    ControlSource ="=[Name]"

                    LayoutCachedLeft =7200
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =240
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

                    LayoutCachedTop =300
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Width =10680
                    TabIndex =5
                    Name ="txtStoreName"
                    Format ="Long Date"

                    LayoutCachedWidth =10680
                    LayoutCachedHeight =240
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =240
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin Line
                    Top =180
                    Width =10680
                    Name ="Line109"
                    LayoutCachedTop =180
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =180
                End
            End
        End
    End
End
CodeBehindForm
' See "ClubPickList.cls"
