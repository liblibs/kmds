Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =0
    TabularFamily =255
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10511
    DatasheetFontHeight =10
    ItemSuffix =55
    Left =735
    Top =5070
    Right =11700
    Bottom =7875
    RecSrcDt = Begin
        0x777287ba7b39e440
    End
    RecordSource ="SELECT ItemMovement.Code, ItemMovement.Description, ItemMovement.CurrMM, ItemMov"
        "ement.YTDQty, ItemMovement.Jan, ItemMovement.Feb, ItemMovement.Mar, ItemMovement"
        ".Apr, ItemMovement.May, ItemMovement.Jun, ItemMovement.Jul, ItemMovement.Aug, It"
        "emMovement.Sep, ItemMovement.Oct, ItemMovement.Nov, ItemMovement.Dec FROM ItemMo"
        "vement ORDER BY ItemMovement.Description; "
    Caption ="qryItemMovementByMonth"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            SpecialEffect =3
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =255
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Width =1560
                    Height =240
                    Name ="Label35"
                    Caption ="TransType"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =3960
                    Width =540
                    Height =240
                    Name ="Label36"
                    Caption ="Jan"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =4500
                    Width =540
                    Height =240
                    Name ="Label37"
                    Caption ="Feb"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    SpecialEffect =2
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =5040
                    Width =540
                    Height =240
                    FontSize =9
                    BackColor =8454016
                    ForeColor =-2147483640
                    Name ="Label38"
                    Caption ="Mar"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5040
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =240
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =5580
                    Width =540
                    Height =240
                    Name ="Label39"
                    Caption ="Apr"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =6120
                    Width =540
                    Height =240
                    Name ="Label40"
                    Caption ="May"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =6660
                    Width =540
                    Height =240
                    Name ="Label41"
                    Caption ="Jun"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =7200
                    Width =540
                    Height =240
                    Name ="Label42"
                    Caption ="Jul"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =7740
                    Width =540
                    Height =240
                    Name ="Label43"
                    Caption ="Aug"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =8280
                    Width =540
                    Height =240
                    Name ="Label44"
                    Caption ="Sep"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =8820
                    Width =540
                    Height =240
                    Name ="Label45"
                    Caption ="Oct"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =9360
                    Width =600
                    Height =240
                    Name ="Label46"
                    Caption ="Nov"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =9960
                    Width =450
                    Height =240
                    Name ="Label47"
                    Caption ="Dec"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =1680
                    Width =900
                    Height =240
                    Name ="lblCurrMM"
                    Caption ="Dec-2011"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =2625
                    Width =735
                    Height =240
                    Name ="Label49"
                    Caption ="Total"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =3345
                    Width =615
                    Height =255
                    Name ="Label50"
                    Caption ="Avg"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =255
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Width =1560
                    ColumnWidth =1500
                    FontSize =10
                    Name ="TransType"
                    ControlSource ="Description"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =2640
                    Width =720
                    Height =255
                    ColumnWidth =840
                    FontSize =10
                    TabIndex =2
                    Name ="YTDQty"
                    ControlSource ="YTDQty"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    Left =3975
                    Width =540
                    Height =255
                    ColumnWidth =495
                    FontSize =9
                    TabIndex =3
                    BackColor =8454016
                    Name ="M1"
                    ControlSource ="Jan"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Here to see the source transaction"

                    LayoutCachedLeft =3975
                    LayoutCachedWidth =4515
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =4515
                    Width =540
                    Height =255
                    ColumnWidth =510
                    FontSize =9
                    TabIndex =4
                    BackColor =8454016
                    Name ="M2"
                    ControlSource ="Feb"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Here to see the source transaction"

                    LayoutCachedLeft =4515
                    LayoutCachedWidth =5055
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =5055
                    Width =540
                    Height =255
                    ColumnWidth =510
                    FontSize =9
                    TabIndex =5
                    BackColor =8454016
                    Name ="M3"
                    ControlSource ="Mar"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Here to see the source transaction"

                    LayoutCachedLeft =5055
                    LayoutCachedWidth =5595
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =5595
                    Width =540
                    Height =255
                    ColumnWidth =480
                    FontSize =9
                    TabIndex =6
                    BackColor =8454016
                    Name ="M4"
                    ControlSource ="Apr"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Here to see the source transaction"

                    LayoutCachedLeft =5595
                    LayoutCachedWidth =6135
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =6135
                    Width =540
                    Height =255
                    ColumnWidth =540
                    FontSize =9
                    TabIndex =7
                    BackColor =8454016
                    Name ="M5"
                    ControlSource ="May"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Here to see the source transaction"

                    LayoutCachedLeft =6135
                    LayoutCachedWidth =6675
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =6675
                    Width =540
                    Height =255
                    ColumnWidth =615
                    FontSize =9
                    TabIndex =8
                    BackColor =8454016
                    Name ="M6"
                    ControlSource ="Jun"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Here to see the source transaction"

                    LayoutCachedLeft =6675
                    LayoutCachedWidth =7215
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =7215
                    Width =540
                    Height =255
                    ColumnWidth =570
                    FontSize =9
                    TabIndex =9
                    BackColor =8454016
                    Name ="M7"
                    ControlSource ="Jul"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Here to see the source transaction"

                    LayoutCachedLeft =7215
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    Left =7740
                    Width =540
                    Height =255
                    ColumnWidth =630
                    FontSize =9
                    TabIndex =10
                    BackColor =8454016
                    Name ="M8"
                    ControlSource ="Aug"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Here to see the source transaction"

                    LayoutCachedLeft =7740
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    Left =8325
                    Width =540
                    Height =255
                    ColumnWidth =600
                    FontSize =9
                    TabIndex =11
                    BackColor =8454016
                    Name ="M9"
                    ControlSource ="Sep"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Here to see the source transaction"

                    LayoutCachedLeft =8325
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =8865
                    Width =540
                    Height =255
                    ColumnWidth =555
                    FontSize =9
                    TabIndex =12
                    BackColor =8454016
                    Name ="M10"
                    ControlSource ="Oct"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Here to see the source transaction"

                    LayoutCachedLeft =8865
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    Left =9405
                    Width =540
                    Height =255
                    ColumnWidth =630
                    FontSize =9
                    TabIndex =13
                    BackColor =8454016
                    Name ="M11"
                    ControlSource ="Nov"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Here to see the source transaction"

                    LayoutCachedLeft =9405
                    LayoutCachedWidth =9945
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =2
                    Left =9945
                    Width =540
                    Height =255
                    ColumnWidth =645
                    FontSize =9
                    TabIndex =14
                    BackColor =8454143
                    Name ="M12"
                    ControlSource ="Dec"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Here to see the source transaction"

                    LayoutCachedLeft =9945
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =1680
                    Width =900
                    Height =255
                    ColumnWidth =795
                    FontSize =10
                    TabIndex =1
                    Name ="CurrMM"
                    ControlSource ="CurrMM"
                    FontName ="Trebuchet MS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Here to see the source transaction"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =3375
                    Width =585
                    Height =255
                    FontSize =10
                    TabIndex =15
                    Name ="Text51"
                    ControlSource ="=[YTDQty]\\12"
                    FontName ="Trebuchet MS"

                End
            End
        End
        Begin FormFooter
            Visible = NotDefault
            Height =660
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =1380
                    Width =960
                    BackColor =8454016
                    Name ="txtCurYear"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    OverlapFlags =85
                    Left =2700
                    Width =960
                    TabIndex =1
                    BackColor =8454143
                    Name ="txtPrevYear"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    Left =180
                    Width =840
                    Height =255
                    ColumnWidth =0
                    TabIndex =2
                    Name ="Code"
                    ControlSource ="Code"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4200
                    Width =1200
                    TabIndex =3
                    Name ="txtDateBegin"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5640
                    Width =1200
                    TabIndex =4
                    Name ="txtDateEnd"

                End
            End
        End
    End
End
CodeBehindForm
' See "PurchaseOrdersOnFileSubItemMovement.cls"
