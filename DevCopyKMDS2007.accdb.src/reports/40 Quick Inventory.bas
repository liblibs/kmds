Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =255
    DateGrouping =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10725
    DatasheetFontHeight =10
    ItemSuffix =38
    Left =525
    Top =1230
    RecSrcDt = Begin
        0x6fbf0680363de440
    End
    RecordSource ="SELECT vwItemsWithCurrentPricing.SalesDepartmentID, vwItemsWithCurrentPricing.ID"
        ", vwItemsWithCurrentPricing.Onhand, IIf([OnShelf]<>[Onhand],[Onshelf],0) AS OnTh"
        "eShelf, vwItemsWithCurrentPricing.Description FROM vwItemsWithCurrentPricing WHE"
        "RE (((vwItemsWithCurrentPricing.SalesDepartmentID) Between Forms!PrintInventoryR"
        "eport!cboSalesDepartmentIDBegin And Forms!PrintInventoryReport!cboSalesDepartmen"
        "tIDEnd) And ((vwItemsWithCurrentPricing.Onhand)<>0) And ((IIf(Forms!PrintInvento"
        "ryReport!chkPrintZeroOnhands,-1,0))=0)) Or (((vwItemsWithCurrentPricing.SalesDep"
        "artmentID) Between Forms!PrintInventoryReport!cboSalesDepartmentIDBegin And Form"
        "s!PrintInventoryReport!cboSalesDepartmentIDEnd) And ((IIf(Forms!PrintInventoryRe"
        "port!chkPrintZeroOnhands,-1,0))=-1)) ORDER BY vwItemsWithCurrentPricing.ID; "
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin PageHeader
            Height =1200
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =60
                    Top =600
                    Width =960
                    Height =240
                    Name ="Label4"
                    Caption ="Item "
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =60
                    LayoutCachedTop =600
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =840
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =1020
                    Top =600
                    Width =465
                    Height =240
                    Name ="Label5"
                    Caption ="Onh"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =1020
                    LayoutCachedTop =600
                    LayoutCachedWidth =1485
                    LayoutCachedHeight =840
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =1485
                    Top =600
                    Width =420
                    Height =240
                    Name ="Label6"
                    Caption ="Shelf"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =1485
                    LayoutCachedTop =600
                    LayoutCachedWidth =1905
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =60
                    Top =300
                    Width =1320
                    ColumnOrder =0
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text48"
                    ControlSource ="=Now()"
                    Format ="Long Time"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =0
                    Left =5700
                    Top =300
                    Width =4365
                    Height =270
                    ColumnOrder =1
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =60
                    Width =3780
                    ColumnOrder =2
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text63"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =0
                    Left =5700
                    Width =4365
                    Height =270
                    ColumnOrder =3
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text64"
                    ControlSource ="=[Name]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Top =300
                    Width =10080
                    ColumnOrder =4
                    FontSize =6
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="Text65"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Width =10020
                    Height =225
                    ColumnOrder =5
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"

                End
                Begin Line
                    LineSlant = NotDefault
                    Top =900
                    Width =10080
                    Name ="Line7"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =2220
                    Top =600
                    Width =960
                    Height =240
                    Name ="Label26"
                    Caption ="Item "
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =2220
                    LayoutCachedTop =600
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =840
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =3180
                    Top =600
                    Width =465
                    Height =240
                    Name ="Label27"
                    Caption ="Onh"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =3180
                    LayoutCachedTop =600
                    LayoutCachedWidth =3645
                    LayoutCachedHeight =840
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =3645
                    Top =600
                    Width =420
                    Height =240
                    Name ="Label28"
                    Caption ="Shelf"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =3645
                    LayoutCachedTop =600
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =840
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =4380
                    Top =600
                    Width =960
                    Height =240
                    Name ="Label29"
                    Caption ="Item "
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =4380
                    LayoutCachedTop =600
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =840
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =5340
                    Top =600
                    Width =465
                    Height =240
                    Name ="Label30"
                    Caption ="Onh"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =5340
                    LayoutCachedTop =600
                    LayoutCachedWidth =5805
                    LayoutCachedHeight =840
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =5805
                    Top =600
                    Width =420
                    Height =240
                    Name ="Label31"
                    Caption ="Shelf"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =5805
                    LayoutCachedTop =600
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =840
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =6540
                    Top =600
                    Width =960
                    Height =240
                    Name ="Label32"
                    Caption ="Item "
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =6540
                    LayoutCachedTop =600
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =840
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =7500
                    Top =600
                    Width =465
                    Height =240
                    Name ="Label33"
                    Caption ="Onh"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =7500
                    LayoutCachedTop =600
                    LayoutCachedWidth =7965
                    LayoutCachedHeight =840
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =7965
                    Top =600
                    Width =420
                    Height =240
                    Name ="Label34"
                    Caption ="Shelf"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =7965
                    LayoutCachedTop =600
                    LayoutCachedWidth =8385
                    LayoutCachedHeight =840
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =8640
                    Top =600
                    Width =960
                    Height =240
                    Name ="Label35"
                    Caption ="Item "
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =8640
                    LayoutCachedTop =600
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =840
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =9600
                    Top =600
                    Width =465
                    Height =240
                    Name ="Label36"
                    Caption ="Onh"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =9600
                    LayoutCachedTop =600
                    LayoutCachedWidth =10065
                    LayoutCachedHeight =840
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =10065
                    Top =600
                    Width =420
                    Height =240
                    Name ="Label37"
                    Caption ="Shelf"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =10065
                    LayoutCachedTop =600
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =840
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Width =1080
                    Height =225
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedWidth =1080
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1560
                    Width =360
                    Height =225
                    TabIndex =1
                    Name ="OnShelf"
                    ControlSource ="OnTheShelf"
                    Format ="#,##0;-#,##0[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1560
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1080
                    Width =420
                    Height =225
                    TabIndex =2
                    Name ="Onhand"
                    ControlSource ="Onhand"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1080
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =225
                End
                Begin Line
                    Left =1980
                    Width =0
                    Height =240
                    Name ="Line3"
                    LayoutCachedLeft =1980
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =240
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
CodeBehindForm
' See "40 Quick Inventory.cls"
