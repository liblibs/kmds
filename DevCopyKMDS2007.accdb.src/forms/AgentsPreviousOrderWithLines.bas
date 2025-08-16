Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =48
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8580
    DatasheetFontHeight =10
    ItemSuffix =21
    Left =4230
    Top =1575
    Right =14685
    Bottom =8820
    Filter ="[InvoHeadID]=3036629"
    RecSrcDt = Begin
        0x76665a43b342e340
    End
    RecordSource ="qryInvolineTransfersByAgentByLineID"
    Caption ="qryInvolineTransfersByAgentByLineID"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
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
            Height =1695
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Width =900
                    Height =255
                    ColumnWidth =900
                    ColumnOrder =0
                    Name ="Store Number"
                    ControlSource ="Store Number"
                    EventProcPrefix ="Store_Number"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Width =1560
                            Height =255
                            Name ="Store Number_Label"
                            Caption ="Store Number"
                            EventProcPrefix ="Store_Number_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =360
                    Width =1035
                    Height =255
                    ColumnWidth =1035
                    ColumnOrder =1
                    TabIndex =1
                    Name ="PostDate"
                    ControlSource ="PostDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =360
                            Width =1560
                            Height =255
                            Name ="PostDate_Label"
                            Caption ="PostDate"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =720
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =2
                    TabIndex =2
                    Name ="StoreName"
                    ControlSource ="StoreName"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =720
                            Width =1560
                            Height =255
                            Name ="StoreName_Label"
                            Caption ="StoreName"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =1080
                    Width =900
                    Height =255
                    ColumnWidth =900
                    ColumnOrder =3
                    TabIndex =3
                    Name ="TransferNumber"
                    ControlSource ="TransferNumber"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =1080
                            Width =1560
                            Height =255
                            Name ="TransferNumber_Label"
                            Caption ="TransferNumber"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =1440
                    Width =900
                    Height =255
                    ColumnWidth =900
                    ColumnOrder =4
                    TabIndex =4
                    Name ="InvoHeadID"
                    ControlSource ="InvoHeadID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =1440
                            Width =1560
                            Height =255
                            Name ="InvoHeadID_Label"
                            Caption ="InvoHeadID"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =3735
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1740
                    Top =120
                    Width =900
                    Height =255
                    ColumnWidth =900
                    Name ="ID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
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
                    Left =1740
                    Top =480
                    Width =2100
                    Height =255
                    ColumnWidth =900
                    TabIndex =1
                    Name ="Code"
                    ControlSource ="Code"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1740
                    LayoutCachedTop =480
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =735
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =480
                            Width =1560
                            Height =255
                            Name ="Code_Label"
                            Caption ="Code"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =2640
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =3
                    Name ="Quantity"
                    ControlSource ="Quantity"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2640
                            Width =1560
                            Height =255
                            Name ="Quantity_Label"
                            Caption ="Quantity"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =3000
                    Width =2310
                    Height =255
                    ColumnWidth =1035
                    TabIndex =4
                    Name ="Cost"
                    ControlSource ="Cost"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3000
                            Width =1560
                            Height =255
                            Name ="Cost_Label"
                            Caption ="Cost"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =3360
                    Width =2310
                    Height =255
                    ColumnWidth =1035
                    TabIndex =5
                    Name ="Price"
                    ControlSource ="Price"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3360
                            Width =1560
                            Height =255
                            Name ="Price_Label"
                            Caption ="Price"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2220
                    Top =1320
                    ColumnWidth =3165
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="Description"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =780
                            Top =1320
                            Width =915
                            Height =240
                            Name ="Label20"
                            Caption ="Description:"
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
' See "AgentsPreviousOrderWithLines.cls"
