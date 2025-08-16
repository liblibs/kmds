Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =48
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5400
    DatasheetFontHeight =10
    ItemSuffix =20
    Left =5190
    Top =4215
    Right =13845
    Bottom =7950
    RecSrcDt = Begin
        0xbf2b82efb242e340
    End
    RecordSource ="qryInvolineTransfersByAgent"
    Caption ="AgentsPreviousOrdersSub"
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
            Height =696
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1740
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
                            Left =120
                            Width =1560
                            Height =255
                            Name ="Store Number_Label"
                            Caption ="Store Number"
                            EventProcPrefix ="Store_Number_Label"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3960
                    Top =120
                    Width =576
                    Height =576
                    TabIndex =1
                    Name ="cmdShowLines"
                    Caption ="Command19"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadad000000000000000a ,
                        0x088888888888880d088888800000080a08888880ffff080d080000800000080a ,
                        0x088888888888880d088888800000080a08888880ffff080d080000800000080a ,
                        0x088888888888880d044444444444440a044444444444440d000000000000000a ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Open Form"

                End
            End
        End
        Begin Section
            Height =4380
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1680
                    Top =480
                    Width =1035
                    Height =255
                    ColumnWidth =1335
                    ColumnOrder =1
                    TabIndex =1
                    Name ="PostDate"
                    ControlSource ="PostDate"
                    Format ="Short Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =1560
                            Height =255
                            Name ="PostDate_Label"
                            Caption ="Date"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1620
                    Width =900
                    Height =255
                    ColumnWidth =825
                    ColumnOrder =3
                    TabIndex =3
                    Name ="TransferNumber"
                    ControlSource ="TransferNumber"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1620
                            Width =1560
                            Height =255
                            Name ="TransferNumber_Label"
                            Caption ="Trans #"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1740
                    Top =120
                    Width =900
                    Height =255
                    ColumnWidth =1140
                    ColumnOrder =0
                    Name ="InvoHeadID"
                    ControlSource ="InvoHeadID"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =1560
                            Height =255
                            Name ="InvoHeadID_Label"
                            Caption ="InvoHeadID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =2340
                    Width =2310
                    Height =255
                    ColumnWidth =915
                    ColumnOrder =6
                    TabIndex =4
                    Name ="TotalQty"
                    ControlSource ="TotalQty"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2340
                            Width =1560
                            Height =255
                            Name ="TotalQty_Label"
                            Caption ="TotalQty"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =2700
                    Width =2310
                    Height =255
                    ColumnWidth =1200
                    ColumnOrder =7
                    TabIndex =5
                    Name ="TotalPrice"
                    ControlSource ="TotalPrice"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2700
                            Width =1560
                            Height =255
                            Name ="TotalPrice_Label"
                            Caption ="TotalPrice"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =3060
                    Width =2310
                    Height =255
                    ColumnWidth =1200
                    ColumnOrder =8
                    TabIndex =6
                    Name ="TotalCost"
                    ControlSource ="TotalCost"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3060
                            Width =1560
                            Height =255
                            Name ="TotalCost_Label"
                            Caption ="TotalCost"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1680
                    Top =855
                    Width =1035
                    Height =255
                    ColumnWidth =1020
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text16"
                    ControlSource ="PostDate"
                    Format ="Medium Time"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =855
                            Width =1560
                            Height =255
                            Name ="Label17"
                            Caption ="Time"
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
' See "AgentsPreviousOrdersSub.cls"
