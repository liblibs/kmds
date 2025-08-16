Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6300
    DatasheetFontHeight =11
    ItemSuffix =6
    Left =6255
    Top =1215
    Right =12810
    Bottom =7200
    RecSrcDt = Begin
        0x00bc534e97f8e340
    End
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Section
            Height =6000
            Name ="Detail"
            Begin
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =720
                    Top =960
                    Width =5160
                    Height =4740
                    Name ="lstSBLC"
                    RowSourceType ="Value List"
                    ColumnWidths ="1440"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =720
                    LayoutCachedTop =960
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =5700
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =720
                            Top =180
                            Width =3480
                            Height =720
                            Name ="lstSBLC_Label"
                            Caption ="These are the batches from today... Double click on the one to restore!"
                            LayoutCachedLeft =720
                            LayoutCachedTop =180
                            LayoutCachedWidth =4200
                            LayoutCachedHeight =900
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =223
                    Left =180
                    Top =120
                    Width =6120
                    Height =5880
                    Name ="Box5"
                    LayoutCachedLeft =180
                    LayoutCachedTop =120
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =6000
                End
            End
        End
    End
End
CodeBehindForm
' See "ScanBatchRestore.cls"
