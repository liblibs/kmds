Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =41
    RecSrcDt = Begin
        0x6789cf4a55bae240
    End
    RecordSource ="qryInvoLineReceiptsSelected"
    Caption ="qryInvoLineReceiptsSelected"
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
            GroupFooter = NotDefault
            ControlSource ="CashSysID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="InvoHeadID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =930
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =4860
                    Height =510
                    FontSize =20
                    Name ="Label34"
                    Caption ="qryInvoLineReceiptsSelected"
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =10800
                    BorderColor =12632256
                    Name ="Line37"
                End
            End
        End
        Begin PageHeader
            Height =435
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =1620
                    Height =300
                    Name ="CashSysID_Label"
                    Caption ="CashSysID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =1740
                    Top =60
                    Width =1440
                    Height =300
                    Name ="InvoHeadID_Label"
                    Caption ="InvoHeadID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3240
                    Top =60
                    Width =600
                    Height =300
                    Name ="ID_Label"
                    Caption ="ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =3900
                    Top =60
                    Width =1200
                    Height =300
                    Name ="Trans_Label"
                    Caption ="Trans"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5160
                    Top =60
                    Width =600
                    Height =300
                    Name ="CashierID_Label"
                    Caption ="CashierID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =5820
                    Top =60
                    Width =1200
                    Height =300
                    Name ="Cashier_Label"
                    Caption ="Cashier"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7080
                    Top =60
                    Width =600
                    Height =300
                    Name ="Code_Label"
                    Caption ="Code"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =7740
                    Top =60
                    Width =1200
                    Height =300
                    Name ="Description_Label"
                    Caption ="Description"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9000
                    Top =60
                    Width =540
                    Height =300
                    Name ="Quantity_Label"
                    Caption ="Quantity"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9600
                    Top =60
                    Width =540
                    Height =300
                    Name ="Price_Label"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =10200
                    Top =60
                    Width =540
                    Height =300
                    Name ="InputCode_Label"
                    Caption ="InputCode"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =10800
                    Name ="Line38"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =10800
                    Name ="Line39"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =420
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Width =1620
                    Height =360
                    FontSize =12
                    FontWeight =700
                    BackColor =12632256
                    Name ="CashSysID"
                    ControlSource ="CashSysID"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =300
            BreakLevel =1
            Name ="GroupHeader2"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =900
                    Height =300
                    FontSize =11
                    ForeColor =8388608
                    Name ="InvoHeadID"
                    ControlSource ="InvoHeadID"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =3300
                    Width =1200
                    Name ="Trans"
                    ControlSource ="Trans"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4560
                    Width =600
                    TabIndex =1
                    Name ="CashierID"
                    ControlSource ="CashierID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5220
                    Width =1200
                    TabIndex =2
                    Name ="Cashier"
                    ControlSource ="Cashier"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6480
                    Width =600
                    TabIndex =3
                    Name ="Code"
                    ControlSource ="Code"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7140
                    Width =1200
                    ColumnWidth =3150
                    TabIndex =4
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8400
                    Width =540
                    TabIndex =5
                    Name ="Quantity"
                    ControlSource ="Quantity"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9000
                    Width =540
                    TabIndex =6
                    Name ="Price"
                    ControlSource ="Price"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9600
                    Width =540
                    TabIndex =7
                    Name ="InputCode"
                    ControlSource ="InputCode"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2520
                    Width =600
                    TabIndex =8
                    Name ="ID"
                    ControlSource ="ID"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =120
            BreakLevel =1
            Name ="GroupFooter3"
            Begin
                Begin Line
                    Left =6300
                    Top =60
                    Width =3780
                    Name ="Line40"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =600
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =540
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Top =240
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text35"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5700
                    Top =240
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text36"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =360
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =9000
                    Width =540
                    Height =300
                    FontSize =10
                    Name ="Quantity Grand Total Sum"
                    ControlSource ="=Sum([Quantity])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Quantity_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label27"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =9600
                    Width =540
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Price Grand Total Sum"
                    ControlSource ="=Sum([Price])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Price_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label31"
                    Caption ="Grand Total"
                End
            End
        End
    End
End
