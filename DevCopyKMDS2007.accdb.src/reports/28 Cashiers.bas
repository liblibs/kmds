Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =8
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10605
    DatasheetFontHeight =10
    ItemSuffix =105
    Left =525
    Top =1230
    RecSrcDt = Begin
        0xd82826f17a70e340
    End
    RecordSource ="qryCashiers"
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
            Height =570
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =30
                    Top =60
                    Width =10470
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Cashiers List"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10500
                    BorderColor =12632256
                    Name ="Line51"
                End
            End
        End
        Begin PageHeader
            Height =540
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =240
                    Top =180
                    Width =420
                    Height =300
                    Name ="Department_Label"
                    Caption ="ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =720
                    Top =180
                    Width =1755
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Name"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =150
                    Width =10455
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =120
                    Width =10455
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Top =480
                    Width =10500
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =315
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin Line
                    Visible = NotDefault
                    Left =180
                    Top =300
                    Width =9960
                    Name ="Line5"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =180
                    Width =480
                    Name ="ID"
                    ControlSource ="ID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =720
                    Width =2220
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Name"

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
                    Width =4785
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
                    Width =4440
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
                    Width =4785
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
                    Width =10500
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
                    Left =180
                    Width =10425
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
' See "28 Cashiers.cls"
