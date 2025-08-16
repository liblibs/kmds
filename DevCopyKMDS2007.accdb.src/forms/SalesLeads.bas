Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    ViewsAllowed =2
    TabularFamily =255
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6300
    RowHeight =330
    DatasheetFontHeight =10
    ItemSuffix =7
    Top =600
    Bottom =600
    RecSrcDt = Begin
        0x5599014f3e09e340
    End
    RecordSource ="SalesLeads"
    Caption ="SalesLeads"
    DatasheetFontName ="Arial"
    AllowFormView =0
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =1560
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =120
                    ColumnWidth =1440
                    Name ="ID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1560
                            Height =255
                            Name ="ID_Label"
                            Caption ="ID"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1620
                    Top =780
                    Width =4560
                    Height =450
                    ColumnWidth =10785
                    TabIndex =2
                    Name ="Text"
                    ControlSource ="Text"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =780
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =780
                            Width =1560
                            Height =255
                            Name ="Text_Label"
                            Caption ="Text"
                            LayoutCachedTop =780
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =1035
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1620
                    Top =450
                    TabIndex =1
                    Name ="chkSpecificItemOnly"
                    ControlSource ="SpecificItemOnly"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =450
                    LayoutCachedWidth =1880
                    LayoutCachedHeight =690
                    DatasheetCaption ="Sales Lead"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1850
                            Top =420
                            Width =630
                            Height =240
                            Name ="Label6"
                            Caption ="Specific Item"
                            LayoutCachedLeft =1850
                            LayoutCachedTop =420
                            LayoutCachedWidth =2480
                            LayoutCachedHeight =660
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
