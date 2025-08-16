Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =127
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10740
    DatasheetFontHeight =10
    ItemSuffix =21
    OrderBy ="qryItemsCartonsToPacks_Step1_Report.PackDesc"
    RecSrcDt = Begin
        0x9d2646618ec8e240
    End
    RecordSource ="qryItemsCartonsToPacks_Step1_Report"
    Caption ="Transfer Cartons To Packs"
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
            ControlSource ="CartonDesc"
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
                    Width =4575
                    Height =510
                    FontSize =20
                    Name ="Label16"
                    Caption ="Transfer Cartons To Packs"
                End
            End
        End
        Begin PageHeader
            Height =780
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =60
                    Top =375
                    Width =765
                    Height =300
                    Name ="CartonDesc_Label"
                    Caption ="Carton "
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =2040
                    Top =375
                    Width =960
                    Height =300
                    Name ="CartonID_Label"
                    Caption ="CartonID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3450
                    Top =120
                    Width =810
                    Height =555
                    Name ="CartonOnhand_Label"
                    Caption ="Carton Onhand"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4620
                    Top =120
                    Width =660
                    Height =555
                    Name ="PacksSold_Label"
                    Caption ="Packs Sold"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5535
                    Top =375
                    Width =765
                    Height =300
                    Name ="PackID_Label"
                    Caption ="PackID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =6360
                    Top =375
                    Width =960
                    Height =300
                    Name ="PackDesc_Label"
                    Caption ="PackDesc"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8490
                    Top =375
                    Width =810
                    Height =300
                    Name ="Onhand_Label"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9705
                    Top =120
                    Width =1035
                    Height =555
                    Name ="PackNewOnhand_Label"
                    Caption ="Pack New Onhand"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =720
                    Width =10680
                    Name ="Line19"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =225
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Width =1920
                    Height =225
                    ColumnWidth =2730
                    Name ="CartonDesc"
                    ControlSource ="CartonDesc"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2040
                    Width =960
                    Height =225
                    TabIndex =1
                    Name ="CartonID"
                    ControlSource ="CartonID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3060
                    Width =1200
                    Height =225
                    ColumnWidth =1425
                    TabIndex =2
                    Name ="CartonOnhand"
                    ControlSource ="CartonOnhand"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4320
                    Width =960
                    Height =225
                    ColumnWidth =1035
                    TabIndex =3
                    Name ="PacksSold"
                    ControlSource ="PacksSold"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5340
                    Width =960
                    Height =225
                    ColumnWidth =810
                    TabIndex =4
                    Name ="PackID"
                    ControlSource ="PackID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6360
                    Width =1920
                    Height =225
                    ColumnWidth =2145
                    TabIndex =5
                    Name ="PackDesc"
                    ControlSource ="PackDesc"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8340
                    Width =960
                    Height =225
                    ColumnWidth =840
                    TabIndex =6
                    Name ="Onhand"
                    ControlSource ="Onhand"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9360
                    Width =1380
                    Height =225
                    ColumnWidth =1605
                    TabIndex =7
                    Name ="PackNewOnhand"
                    ControlSource ="PackNewOnhand"

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
                    Top =300
                    Width =780
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
                    Left =6660
                    Top =240
                    Width =3900
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
                    Width =3780
                    Height =225
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
                    Left =6660
                    Width =3900
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
                    Top =300
                    Width =10620
                    FontSize =6
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
                    Left =180
                    Width =10440
                    Height =225
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
            Height =0
            Name ="ReportFooter"
        End
    End
End
