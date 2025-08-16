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
    ItemSuffix =83
    Left =525
    Top =1230
    OrderBy ="ID"
    RecSrcDt = Begin
        0x37a3d218d36de340
    End
    RecordSource ="qrySizes"
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
                    Left =45
                    Top =60
                    Width =10695
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Size List"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10755
                    BorderColor =12632256
                    Name ="Line51"
                End
            End
        End
        Begin PageHeader
            Height =600
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =240
                    Top =180
                    Width =1155
                    Height =300
                    Name ="Department_Label"
                    Caption ="ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =1920
                    Top =180
                    Width =1755
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Description"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5100
                    Top =180
                    Width =540
                    Height =300
                    Name ="D1_Label"
                    Caption =" OZ."
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6240
                    Top =180
                    Width =930
                    Height =300
                    Name ="D2_Label"
                    Caption ="Case Qty"
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
                    Top =120
                    Width =10710
                    Name ="Line55"
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
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =480
                    Width =10710
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =7440
                    Top =180
                    Width =1440
                    Height =300
                    Name ="Label39"
                    Caption ="SKU Count"
                    LayoutCachedLeft =7440
                    LayoutCachedTop =180
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =480
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin Line
                    Visible = NotDefault
                    Left =180
                    Top =240
                    Width =9960
                    Name ="Line5"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =180
                    Width =1200
                    Name ="ID"
                    ControlSource ="ID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1920
                    Width =2220
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4200
                    TabIndex =2
                    Name ="Oz"
                    ControlSource ="Oz"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5760
                    TabIndex =3
                    Name ="CaseQty"
                    ControlSource ="CaseQty"

                End
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =7920
                    Width =720
                    Height =225
                    TabIndex =4
                    Name ="Text40"
                    ControlSource ="Items"
                    Format ="Standard"

                    LayoutCachedLeft =7920
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =225
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
                    Width =4920
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
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "20 Size List.cls"
