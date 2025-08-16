Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8280
    DatasheetFontHeight =10
    ItemSuffix =11
    RecSrcDt = Begin
        0xd48cd0b07021e340
    End
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
        Begin Section
            Height =12060
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =4140
                    Top =7080
                    Width =1680
                    Height =420
                    Name ="Command2"
                    Caption ="Command2"
                    OnClick ="[Event Procedure]"

                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    TextFontFamily =49
                    MultiSelect =1
                    IMESentenceMode =3
                    Left =420
                    Top =120
                    Width =4260
                    Height =6480
                    TabIndex =1
                    Name ="lstTables"
                    RowSourceType ="Value List"
                    ColumnWidths ="1440"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Courier"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =420
                            Top =120
                            Width =420
                            Height =245
                            Name ="Col1_Label"
                            Caption ="Col1"
                        End
                    End
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextFontFamily =49
                    MultiSelect =1
                    IMESentenceMode =3
                    Left =4920
                    Top =120
                    Width =3360
                    Height =6480
                    TabIndex =2
                    Name ="lstTables1"
                    RowSourceType ="Value List"
                    ColumnWidths ="1440"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Courier"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4020
                            Top =120
                            Width =420
                            Height =245
                            Name ="Label10"
                            Caption ="Col1"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "LocalTablesDelete.cls"
