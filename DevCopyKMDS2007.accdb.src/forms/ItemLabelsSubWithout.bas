Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =119
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8940
    DatasheetFontHeight =10
    ItemSuffix =27
    Left =570
    Right =14340
    Bottom =5205
    Tag ="0"
    Filter ="(( ItemLabels.Labels=0))"
    ShortcutMenuBar ="KMDSPopup"
    OrderBy ="[ItemLabels].[ItemD], [ItemLabels].[ChangeDate] DESC, [ItemLabels].[ID]"
    RecSrcDt = Begin
        0x8702191b96a7e240
    End
    RecordSource ="ItemLabels"
    Caption ="ItemLabels"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Trebuchet MS"
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
            Height =4095
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =480
                    Width =900
                    Height =255
                    ColumnWidth =900
                    ColumnOrder =1
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =1380
                            Height =255
                            Name ="ID_Label"
                            Caption ="ID"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =840
                    Width =2310
                    Height =255
                    ColumnWidth =2130
                    ColumnOrder =2
                    TabIndex =1
                    Name ="ItemD"
                    ControlSource ="ItemD"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =840
                            Width =1380
                            Height =255
                            Name ="ItemD_Label"
                            Caption ="ItemD"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =1200
                    Width =1935
                    Height =255
                    ColumnWidth =1320
                    ColumnOrder =3
                    TabIndex =2
                    Name ="SizeD"
                    ControlSource ="SizeD"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1200
                            Width =1380
                            Height =255
                            Name ="SizeD_Label"
                            Caption ="SizeD"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =1560
                    Width =2310
                    Height =255
                    ColumnWidth =915
                    ColumnOrder =4
                    TabIndex =3
                    Name ="CurPrice"
                    ControlSource ="CurPrice"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1560
                            Width =1380
                            Height =255
                            Name ="CurPrice_Label"
                            Caption ="CurPrice"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =1920
                    Width =4020
                    Height =450
                    ColumnWidth =285
                    ColumnOrder =5
                    TabIndex =4
                    Name ="S"
                    ControlSource ="S"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1920
                            Width =1380
                            Height =255
                            Name ="S_Label"
                            Caption ="S"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =2460
                    Width =2310
                    Height =255
                    ColumnWidth =1185
                    ColumnOrder =6
                    TabIndex =5
                    Name ="SalesDpt"
                    ControlSource ="SalesDpt"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2460
                            Width =1380
                            Height =255
                            Name ="SalesDpt_Label"
                            Caption ="SalesDpt"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =2820
                    Width =1935
                    Height =255
                    ColumnWidth =1620
                    ColumnOrder =7
                    TabIndex =6
                    Name ="VendorName"
                    ControlSource ="VendorName"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2820
                            Width =1380
                            Height =255
                            Name ="VendorName_Label"
                            Caption ="VendorName"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =3180
                    Width =4020
                    ColumnWidth =1320
                    ColumnOrder =8
                    TabIndex =7
                    Name ="UPCCodeMax"
                    ControlSource ="UPCCodeMax"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3180
                            Width =1380
                            Height =255
                            Name ="UPCCodeMax_Label"
                            Caption ="UPCCodeMax"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =3480
                    Width =1035
                    Height =255
                    ColumnWidth =1125
                    ColumnOrder =9
                    TabIndex =8
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    Format ="Short Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3480
                            Width =1380
                            Height =255
                            Name ="SetupDate_Label"
                            Caption ="SetupDate"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =3780
                    Width =840
                    Height =255
                    ColumnWidth =1230
                    ColumnOrder =10
                    TabIndex =9
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    Format ="Short Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3780
                            Width =1380
                            Height =255
                            Name ="ChangeDate_Label"
                            Caption ="ChangeDate"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =120
                    ColumnOrder =0
                    TabIndex =10
                    Name ="Labels"
                    ControlSource ="Labels"
                    OnDblClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =780
                            Top =120
                            Width =615
                            Height =240
                            Name ="Label26"
                            Caption ="Labels:"
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
CodeBehindForm
' See "ItemLabelsSubWithout.cls"
