Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7920
    DatasheetFontHeight =10
    ItemSuffix =32
    OrderBy ="Agents.City"
    RecSrcDt = Begin
        0x47efb7df71ade240
    End
    RecordSource ="Agents"
    Caption ="Agents1"
    DatasheetFontName ="Arial"
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
            Height =4275
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =120
                    Width =900
                    Height =255
                    ColumnWidth =900
                    Name ="AgentNumber"
                    ControlSource ="AgentNumber"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1200
                            Height =255
                            Name ="AgentNumber_Label"
                            Caption ="AgentNumber"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =480
                    Width =3480
                    Height =450
                    ColumnWidth =3000
                    TabIndex =1
                    Name ="FirstName"
                    ControlSource ="FirstName"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =1200
                            Height =255
                            Name ="FirstName_Label"
                            Caption ="FirstName"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =1020
                    Width =3480
                    Height =450
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="LastName"
                    ControlSource ="LastName"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1020
                            Width =1200
                            Height =255
                            Name ="LastName_Label"
                            Caption ="LastName"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =1560
                    Width =3480
                    Height =450
                    ColumnWidth =3000
                    TabIndex =3
                    Name ="Address"
                    ControlSource ="Address"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1560
                            Width =1200
                            Height =255
                            Name ="Address_Label"
                            Caption ="Address"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =2100
                    Width =3480
                    Height =450
                    ColumnWidth =3000
                    TabIndex =4
                    Name ="City"
                    ControlSource ="City"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2100
                            Width =1200
                            Height =255
                            Name ="City_Label"
                            Caption ="City"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =2640
                    Width =3480
                    Height =450
                    ColumnWidth =3000
                    TabIndex =5
                    Name ="ST"
                    ControlSource ="ST"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2640
                            Width =1200
                            Height =255
                            Name ="ST_Label"
                            Caption ="ST"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =3180
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =6
                    Name ="Zip"
                    ControlSource ="Zip"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3180
                            Width =1200
                            Height =255
                            Name ="Zip_Label"
                            Caption ="Zip"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =3540
                    Width =1185
                    Height =255
                    ColumnWidth =1185
                    TabIndex =7
                    Name ="Phone"
                    ControlSource ="Phone"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3540
                            Width =1200
                            Height =255
                            Name ="Phone_Label"
                            Caption ="Phone"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =3900
                    Width =1185
                    Height =255
                    ColumnWidth =1185
                    TabIndex =8
                    Name ="Fax/Modem"
                    ControlSource ="Fax/Modem"
                    EventProcPrefix ="Fax_Modem"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3900
                            Width =1200
                            Height =255
                            Name ="Fax/Modem_Label"
                            Caption ="Fax/Modem"
                            EventProcPrefix ="Fax_Modem_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6120
                    Top =120
                    Width =1185
                    Height =255
                    ColumnWidth =1185
                    TabIndex =9
                    Name ="Home"
                    ControlSource ="Home"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4860
                            Top =120
                            Width =1200
                            Height =255
                            Name ="Home_Label"
                            Caption ="Home"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6120
                    Top =480
                    Width =1185
                    Height =255
                    ColumnWidth =1185
                    TabIndex =10
                    Name ="Mobile"
                    ControlSource ="Mobile"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4860
                            Top =480
                            Width =1200
                            Height =255
                            Name ="Mobile_Label"
                            Caption ="Mobile"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6120
                    Top =840
                    Width =615
                    Height =255
                    ColumnWidth =615
                    TabIndex =11
                    Name ="ActiveForTransfer"
                    ControlSource ="ActiveForTransfer"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4860
                            Top =840
                            Width =1200
                            Height =255
                            Name ="ActiveForTransfer_Label"
                            Caption ="ActiveForTransfer"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6120
                    Top =1200
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =12
                    Name ="OrderCount"
                    ControlSource ="OrderCount"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4860
                            Top =1200
                            Width =1200
                            Height =255
                            Name ="OrderCount_Label"
                            Caption ="OrderCount"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6120
                    Top =1560
                    Width =1035
                    Height =255
                    ColumnWidth =1035
                    TabIndex =13
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4860
                            Top =1560
                            Width =1200
                            Height =255
                            Name ="SetupDate_Label"
                            Caption ="SetupDate"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6120
                    Top =1920
                    Width =1035
                    Height =255
                    ColumnWidth =1035
                    TabIndex =14
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4860
                            Top =1920
                            Width =1200
                            Height =255
                            Name ="ChangeDate_Label"
                            Caption ="ChangeDate"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6120
                    Top =2280
                    Width =1740
                    Height =255
                    ColumnWidth =2310
                    TabIndex =15
                    Name ="upsize_ts"
                    ControlSource ="upsize_ts"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4860
                            Top =2280
                            Width =1200
                            Height =255
                            Name ="upsize_ts_Label"
                            Caption ="upsize_ts"
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
