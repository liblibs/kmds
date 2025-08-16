Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =8
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =86
    Left =525
    Top =1230
    OrderBy ="CartonDesc"
    RecSrcDt = Begin
        0xfda15e7ef4a8e240
    End
    RecordSource ="qryItemsWithASubPack"
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
            ControlSource ="CartonSD"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1155
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =600
                    Width =10621
                    Height =420
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =10620
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Items with Packs "
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
            Height =480
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =240
                    Top =180
                    Width =1170
                    Height =300
                    Name ="Department_Label"
                    Caption ="Carton Info"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5280
                    Top =180
                    Width =1065
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Pack UPC"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4815
                    Top =180
                    Width =390
                    Height =300
                    Name ="D1_Label"
                    Caption ="Qty"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =6480
                    Top =180
                    Width =1005
                    Height =300
                    Name ="D2_Label"
                    Caption ="Pack Info"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =150
                    Width =10710
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =450
                    Width =10710
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3825
                    Top =180
                    Width =795
                    Height =300
                    Name ="Label84"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9900
                    Top =180
                    Width =795
                    Height =300
                    Name ="Label85"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =375
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =180
                    Width =672
                    Height =375
                    ColumnWidth =1260
                    FontSize =14
                    Name ="Department"
                    ControlSource ="CartonSD"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextFontFamily =18
                    Left =900
                    Width =2304
                    Height =375
                    ColumnWidth =1710
                    FontSize =14
                    TabIndex =1
                    Name ="DepartmentDesc"
                    ControlSource ="CartonSDDesc"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =315
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    Left =60
                    Width =840
                    Height =225
                    ColumnWidth =570
                    TabIndex =1
                    Name ="ID"
                    ControlSource ="CartonID"

                End
                Begin TextBox
                    Left =960
                    Width =2880
                    Height =225
                    ColumnWidth =3120
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="CartonDesc"

                End
                Begin TextBox
                    Left =6120
                    Width =840
                    Height =225
                    TabIndex =3
                    Name ="Text77"
                    ControlSource ="PackID"

                End
                Begin TextBox
                    Left =7020
                    Width =2880
                    Height =225
                    TabIndex =4
                    Name ="Text78"
                    ControlSource ="PackDesc"

                End
                Begin TextBox
                    TextAlign =3
                    Left =5205
                    Width =840
                    Height =225
                    TabIndex =5
                    Name ="Text79"
                    ControlSource ="CartonSUBUpc"

                End
                Begin TextBox
                    TextAlign =3
                    Left =4680
                    Width =420
                    Height =225
                    TabIndex =6
                    Name ="Text80"
                    ControlSource ="CartonPack"

                End
                Begin Line
                    Visible = NotDefault
                    LineSlant = NotDefault
                    Left =60
                    Top =300
                    Width =10740
                    Name ="Line5"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9960
                    Width =780
                    Height =225
                    Name ="Text82"
                    ControlSource ="PackPrice"
                    Format ="Standard"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3840
                    Width =780
                    Height =225
                    TabIndex =7
                    Name ="Text83"
                    ControlSource ="CartonPrice"
                    Format ="Standard"

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
                    Left =120
                    Width =4500
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
                    Name ="Text64"
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
                    Name ="Text65"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =120
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
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "08 Items with Packs.cls"
