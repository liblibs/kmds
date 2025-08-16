Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =119
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9000
    DatasheetFontHeight =10
    ItemSuffix =14
    Top =600
    Bottom =600
    RecSrcDt = Begin
        0x85a8863647afe240
    End
    RecordSource ="Sysfile"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Trebuchet MS"
    OnError ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Tab
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Section
            Height =6600
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =1500
                    Top =5940
                    Width =1740
                    Height =540
                    TabIndex =2
                    Name ="Command6"
                    Caption ="Process files"
                    OnClick ="[Event Procedure]"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1500
                    Top =60
                    Width =1980
                    TabIndex =3
                    Name ="AgentNumber"
                    ControlSource ="AgentNumber"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =180
                            Top =60
                            Width =1260
                            Height =240
                            Name ="Label7"
                            Caption ="AgentNumber:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3540
                    Top =60
                    Width =5040
                    TabIndex =4
                    Name ="Name"
                    ControlSource ="Name"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =420
                    Width =1980
                    TabIndex =5
                    Name ="NetworkMailDir"
                    ControlSource ="NetworkMailDir"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =420
                            Width =1380
                            Height =240
                            Name ="Label9"
                            Caption ="NetworkMailDir:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3540
                    Top =420
                    Width =4980
                    TabIndex =6
                    Name ="NetworkMailDirFileSpec"

                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =360
                    Top =1200
                    Width =4140
                    Height =4500
                    Name ="txtToBeProcessed"
                    RowSourceType ="Value List"
                    BeforeUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =360
                            Top =960
                            Width =1590
                            Height =240
                            Name ="Label3"
                            Caption ="Files to be processed"
                        End
                    End
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =4620
                    Top =1140
                    Width =4140
                    Height =4560
                    TabIndex =1
                    Name ="txtProcessed"
                    RowSourceType ="Value List"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4620
                            Top =900
                            Width =1590
                            Height =240
                            Name ="Label5"
                            Caption ="Processed"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ImportInvoiceLDFromCSV.cls"
