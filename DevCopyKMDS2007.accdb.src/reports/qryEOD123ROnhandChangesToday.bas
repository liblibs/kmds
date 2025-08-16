Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =8
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10770
    DatasheetFontHeight =10
    ItemSuffix =73
    Top =1080
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x4ba6310ccebee240
    End
    RecordSource ="qry EODOnhandChanges"
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
            ControlSource ="SalesDepartment.ID"
        End
        Begin BreakLevel
            ControlSource ="Item.ID"
        End
        Begin BreakLevel
            ControlSource ="SetupDate"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1155
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    Left =60
                    Top =600
                    Width =10695
                    Height =510
                    ColumnOrder =0
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =3015
                    Top =60
                    Width =4815
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Audit Onhand Change"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10755
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10755
                    BorderColor =12632256
                    Name ="Line53"
                End
            End
        End
        Begin PageHeader
            Height =840
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =240
                    Top =120
                    Width =1320
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3735
                    Top =120
                    Width =1365
                    Height =300
                    Name ="D2_Label"
                    Caption ="Qty"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =150
                    Width =10710
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =120
                    Width =10710
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =450
                    Width =10710
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =480
                    Width =10710
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5385
                    Top =120
                    Width =555
                    Height =300
                    Name ="Label57"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6240
                    Top =120
                    Width =960
                    Height =300
                    Name ="Label68"
                    Caption ="ExttPrice"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =510
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Top =120
                    Width =588
                    Height =330
                    FontSize =12
                    FontWeight =700
                    BackColor =12632256
                    Name ="Group"
                    ControlSource ="SalesDepartment.ID"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextFontFamily =18
                    Left =720
                    Top =120
                    Width =1503
                    Height =330
                    ColumnWidth =1185
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    BackColor =12632256
                    Name ="GroupDesc"
                    ControlSource ="SalesDepartment.Description"
                    FontName ="Times New Roman"

                End
                Begin Label
                    Visible = NotDefault
                    Left =3900
                    Top =120
                    Width =3240
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    Left =60
                    Width =1260
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="Item.ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =1440
                    Width =2880
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Item.Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1440
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =4380
                    Width =720
                    Height =225
                    TabIndex =2
                    Name ="Onhand"
                    ControlSource ="Quantity"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5220
                    Width =720
                    Height =225
                    TabIndex =3
                    Name ="Text58"
                    ControlSource ="UPrice"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6120
                    Width =1080
                    TabIndex =4
                    Name ="Ext"
                    ControlSource ="Ext"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7380
                    Width =480
                    TabIndex =5
                    Name ="Cashier.ID"
                    ControlSource ="Cashier.ID"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Cashier_ID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7860
                    TabIndex =6
                    Name ="Name"
                    ControlSource ="Name"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9300
                    ColumnWidth =2100
                    TabIndex =7
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    Format ="Medium Time"
                    FontName ="Arial Narrow"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =450
            Name ="GroupFooter4"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =4380
                    Top =120
                    Width =720
                    Height =225
                    Name ="Text66"
                    ControlSource ="=Sum([Quantity])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6120
                    Top =120
                    Width =1080
                    FontWeight =700
                    TabIndex =1
                    Name ="Text67"
                    ControlSource ="=Sum([Ext])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    Left =2700
                    Top =120
                    Width =1503
                    Height =330
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    BackColor =12632256
                    Name ="Text71"
                    ControlSource ="SalesDepartment.Description"
                    FontName ="Arial Narrow"

                End
            End
        End
        Begin PageFooter
            Height =600
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Top =300
                    Width =5040
                    Height =300
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
                    Height =300
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
                    Height =300
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
                    Height =300
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
                    Left =60
                    Top =300
                    Width =10680
                    Height =300
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
            Height =300
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    Left =4440
                    Top =60
                    Width =720
                    Height =225
                    Name ="Text69"
                    ControlSource ="=Sum([Quantity])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6180
                    Top =60
                    Width =1080
                    TabIndex =1
                    Name ="Text70"
                    ControlSource ="=Sum([Ext])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Width =1320
                    Height =300
                    Name ="Label72"
                    Caption ="Report total"
                    FontName ="Arial Narrow"
                    Tag ="DetachedLabel"
                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD123ROnhandChangesToday.cls"
