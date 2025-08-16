Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =119
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7620
    DatasheetFontHeight =10
    ItemSuffix =9
    RecSrcDt = Begin
        0xf375a3d371abe240
    End
    RecordSource ="UPCItem"
    Caption ="UPCItem"
    DatasheetFontName ="Trebuchet MS"
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
            Height =5340
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =120
                    Width =1935
                    Height =255
                    ColumnWidth =1935
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
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =480
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =1
                    Name ="ItemID"
                    ControlSource ="ItemID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =1560
                            Height =255
                            Name ="ItemID_Label"
                            Caption ="ItemID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =5520
                    Top =180
                    Width =1680
                    Height =720
                    FontSize =24
                    TabIndex =2
                    Name ="txtUPCA"
                    FontName ="UPC-A"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4080
                            Top =180
                            Width =555
                            Height =240
                            Name ="Label5"
                            Caption ="UPC-A"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2280
                    Top =3120
                    Width =1620
                    Height =780
                    FontSize =24
                    TabIndex =3
                    Name ="txtUPCE"
                    FontName ="UPC-E"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =840
                            Top =3120
                            Width =555
                            Height =240
                            Name ="Label7"
                            Caption ="UPC-e"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1680
                    Top =960
                    TabIndex =4
                    Name ="Command8"
                    Caption ="Command8"
                    OnClick ="[Event Procedure]"

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
' See "UPCItem.cls"
