Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =4
    TabularFamily =119
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =30870
    DatasheetFontHeight =10
    ItemSuffix =33
    RecSrcDt = Begin
        0x0597170fd1aae240
    End
    RecordSource ="qryItemChartMovement"
    Caption ="qryItemChartMovement"
    DatasheetFontName ="Trebuchet MS"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:PivotTable>\015"
        "\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  <x:NoDisplayAle"
        "rts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTips/>\015\012  <x"
        ":NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>541</x:Width>\015"
        "\012  <x:NoAllowGrouping/>\015\012  <x:NoAllowFiltering/>\015\012  <x:CubeProvid"
        "er>msolap.2</x:CubeProvider>\015\012  <x:CacheDetails/>\015\012  <x:Name>Microso"
        "ft Office PivotTable 10.0</x:Name>\015\012  <x:PivotField>\015\012   <x:Name>Cod"
        "e</x:Name>\015\012   <x:DataType>Integer</x:DataType>\015\012   <x:EncodedType>a"
        "dInteger</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Comp"
        "areOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   "
        "<x:Name>Description</x:Name>\015\012   <x:Orientation>Page</x:Orientation>\015\012"
        "   <x:EncodedType>adLongVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMemb"
        "ersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <"
        "x:PivotField>\015\012   <x:Name>Onhand</x:Name>\015\012   <x:DataType>Integer</x"
        ":DataType>\015\012   <x:PLDataOrientation/>\015\012   <x:EncodedType>adInteger</"
        "x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrdered"
        "MembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Mi"
        "n</x:Name>\015\012   <x:SourceName>OrderAt</x:SourceName>\015\012   <x:DataType>"
        "Integer</x:DataType>\015\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>2<"
        "/x:PLPosition>\015\012   <x:EncodedType>adInteger</x:EncodedType>\015\012   <x:C"
        "ompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:Pivot"
        "Field>\015\012  <x:PivotField>\015\012   <x:Name>Max</x:Name>\015\012   <x:Sourc"
        "eName>OrderTo</x:SourceName>\015\012   <x:DataType>Integer</x:DataType>\015\012 "
        "  <x:PLDataOrientation/>\015\012   <x:PLPosition>3</x:PLPosition>\015\012   <x:E"
        "ncodedType>adInteger</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Unique"
        "Name</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField"
        ">\015\012   <x:Name>CurrMM</x:Name>\015\012   <x:DataType>Integer</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>4</x:PLPosition>\015\012  "
        " <x:EncodedType>adInteger</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>U"
        "niqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>Jan</x:Name>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>5</x:PLPosition>\015\012  "
        " <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Un"
        "iqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotF"
        "ield>\015\012   <x:Name>Feb</x:Name>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>6</x:PLPosition>\015\012  "
        " <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Un"
        "iqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotF"
        "ield>\015\012   <x:Name>Mar</x:Name>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>7</x:PLPosition>\015\012  "
        " <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Un"
        "iqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotF"
        "ield>\015\012   <x:Name>Apr</x:Name>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>8</x:PLPosition>\015\012  "
        " <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Un"
        "iqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotF"
        "ield>\015\012   <x:Name>May</x:Name>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>9</x:PLPosition>\015\012  "
        " <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Un"
        "iqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotF"
        "ield>\015\012   <x:Name>Jun</x:Name>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>10</x:PLPosition>\015\012 "
        "  <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>U"
        "niqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>Jul</x:Name>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>11</x:PLPosition>\015\012 "
        "  <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>U"
        "niqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>Aug</x:Name>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>12</x:PLPosition>\015\012 "
        "  <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>U"
        "niqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>Sep</x:Name>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>13</x:PLPosition>\015\012 "
        "  <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>U"
        "niqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>Oct</x:Name>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>14</x:PLPosition>\015\012 "
        "  <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>U"
        "niqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>Nov</x:Name>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>15</x:PLPosition>\015\012 "
        "  <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>U"
        "niqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>Dec</x:Name>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>16</x:PLPosition>\015\012 "
        "  <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>U"
        "niqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>Data</x:Name>\015\012   <x:Orientation>Column</x:Orient"
        "ation>\015\012   <x:Position>-1</x:Position>\015\012   <x:DataField/>\015\012  <"
        "/x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Sum of Onhand</x:Name>\015"
        "\012   <x:PLName>Total1</x:PLName>\015\012   <x:TotalNumber>0</x:TotalNumber>\015"
        "\012   <x:Orientation>Data</x:Orientation>\015\012   <x:Position>1</x:Position>\015"
        "\012   <x:ParentField>Onhand</x:ParentField>\015\012  </x:PivotField>\015\012  <"
        "x:PivotField>\015\012   <x:Name>Sum of Min</x:Name>\015\012   <x:PLName>Total2</"
        "x:PLName>\015\012   <x:TotalNumber>1</x:TotalNumber>\015\012   <x:Orientation>Da"
        "ta</x:Orientation>\015\012   <x:Position>2</x:Position>\015\012   <x:ParentField"
        ">Min</x:ParentField>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   "
        "<x:Name>Sum of Max</x:Name>\015\012   <x:PLName>Total3</x:PLName>\015\012   <x:T"
        "otalNumber>2</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation>\015\012"
        "   <x:Position>3</x:Position>\015\012   <x:ParentField>Max</x:ParentField>\015\012"
        "  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Sum of CurrMM</x:Nam"
        "e>\015\012   <x:PLName>Total4</x:PLName>\015\012   <x:TotalNumber>3</x:TotalNumb"
        "er>\015\012   <x:Orientation>Data</x:Orientation>\015\012   <x:Position>4</x:Pos"
        "ition>\015\012   <x:ParentField>CurrMM</x:ParentField>\015\012  </x:PivotField>\015"
        "\012  <x:PivotField>\015\012   <x:Name>Sum of Jan</x:Name>\015\012   <x:PLName>T"
        "otal5</x:PLName>\015\012   <x:TotalNumber>4</x:TotalNumber>\015\012   <x:Orienta"
        "tion>Data</x:Orientation>\015\012   <x:Position>5</x:Position>\015\012   <x:Pare"
        "ntField>Jan</x:ParentField>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>Sum of Feb</x:Name>\015\012   <x:PLName>Total6</x:PLName>\015\012"
        "   <x:TotalNumber>5</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation"
        ">\015\012   <x:Position>6</x:Position>\015\012   <x:ParentField>Feb</x:ParentFie"
        "ld>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Sum of Ma"
        "r</x:Name>\015\012   <x:PLName>Total7</x:PLName>\015\012   <x:TotalNumber>6</x:T"
        "otalNumber>\015\012   <x:Orientation>Data</x:Orientation>\015\012   <x:Position>"
        "7</x:Position>\015\012   <x:ParentField>Mar</x:ParentField>\015\012  </x:PivotFi"
        "eld>\015\012  <x:PivotField>\015\012   <x:Name>Sum of Apr</x:Name>\015\012   <x:"
        "PLName>Total8</x:PLName>\015\012   <x:TotalNumber>7</x:TotalNumber>\015\012   <x"
        ":Orientation>Data</x:Orientation>\015\012   <x:Position>8</x:Position>\015\012  "
        " <x:ParentField>Apr</x:ParentField>\015\012  </x:PivotField>\015\012  <x:PivotFi"
        "eld>\015\012   <x:Name>Sum of May</x:Name>\015\012   <x:PLName>Total9</x:PLName>"
        "\015\012   <x:TotalNumber>8</x:TotalNumber>\015\012   <x:Orientation>Data</x:Ori"
        "entation>\015\012   <x:Position>9</x:Position>\015\012   <x:ParentField>May</x:P"
        "arentField>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>S"
        "um of Jun</x:Name>\015\012   <x:PLName>Total10</x:PLName>\015\012   <x:TotalNumb"
        "er>9</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation>\015\012   <x:"
        "Position>10</x:Position>\015\012   <x:ParentField>Jun</x:ParentField>\015\012  <"
        "/x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Sum of Jul</x:Name>\015"
        "\012   <x:PLName>Total11</x:PLName>\015\012   <x:TotalNumber>10</x:TotalNumber>\015"
        "\012   <x:Orientation>Data</x:Orientation>\015\012   <x:Position>11</x:Position>"
        "\015\012   <x:ParentField>Jul</x:ParentField>\015\012  </x:PivotField>\015\012  "
        "<x:PivotField>\015\012   <x:Name>Sum of Aug</x:Name>\015\012   <x:PLName>Total12"
        "</x:PLName>\015\012   <x:TotalNumber>11</x:TotalNumber>\015\012   <x:Orientation"
        ">Data</x:Orientation>\015\012   <x:Position>12</x:Position>\015\012   <x:ParentF"
        "ield>Aug</x:ParentField>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012"
        "   <x:Name>Sum of Sep</x:Name>\015\012   <x:PLName>Total13</x:PLName>\015\012   "
        "<x:TotalNumber>12</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation>\015"
        "\012   <x:Position>13</x:Position>\015\012   <x:ParentField>Sep</x:ParentField>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Sum of Oct</x:Na"
        "me>\015\012   <x:PLName>Total14</x:PLName>\015\012   <x:TotalNumber>13</x:TotalN"
        "umber>\015\012   <x:Orientation>Data</x:Orientation>\015\012   <x:Position>14</x"
        ":Position>\015\012   <x:ParentField>Oct</x:ParentField>\015\012  </x:PivotField>"
        "\015\012  <x:PivotField>\015\012   <x:Name>Sum of Nov</x:Name>\015\012   <x:PLNa"
        "me>Total15</x:PLName>\015\012   <x:TotalNumber>14</x:TotalNumber>\015\012   <x:O"
        "rientation>Data</x:Orientation>\015\012   <x:Position>15</x:Position>\015\012   "
        "<x:ParentField>Nov</x:ParentField>\015\012  </x:PivotField>\015\012  <x:PivotFie"
        "ld>\015\012   <x:Name>Sum of Dec</x:Name>\015\012   <x:PLName>Total16</x:PLName>"
        "\015\012   <x:TotalNumber>15</x:TotalNumber>\015\012   <x:Orientation>Data</x:Or"
        "ientation>\015\012   <x:Position>16</x:Position>\015\012   <x:ParentField>Dec</x"
        ":ParentField>\015\012  </x:PivotField>\015\012  <x:PivotData>\015\012   <x:Top>0"
        "</x:Top>\015\012   <x:TopOffset>0</x:TopOffset>\015\012   <x:Left>0</x:Left>\015"
        "\012   <x:LeftOffset>0</x:LeftOffset>\015\012   <x:SeqNum>364</x:SeqNum>\015\012"
        "  </x:PivotData>\015\012  <x:PivotView>\015\012   <x:Label>\015\012    <x:Captio"
        "n>Microsoft Office PivotTable 10.0</x:Caption>\015\012    <x:NotVisible/>\015\012"
        "   </x:Label>\015\012   <x:AllowEdits>false</x:AllowEdits>\015\012   <x:AllowAdd"
        "itions>false</x:AllowAdditions>\015\012   <x:AllowDeletions>false</x:AllowDeleti"
        "ons>\015\012  </x:PivotView>\015\012 </x:PivotTable>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  <x:Width>22172<"
        "/x:Width>\015\012  <x:Height>3863</x:Height>\015\012  <x:AllowPropertyBrowser/>\015"
        "\012  <x:DataSource>\015\012   <x:Type>PivotList</x:Type>\015\012  </x:DataSourc"
        "e>\015\012  <x:BoundSeries>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015"
        "\012  </x:BoundSeries>\015\012  <x:Category>\015\012   <x:DataSourceIndex>0</x:D"
        "ataSourceIndex>\015\012  </x:Category>\015\012  <x:Value>\015\012   <x:DataSourc"
        "eIndex>0</x:DataSourceIndex>\015\012  </x:Value>\015\012  <x:BoundCharts>\015\012"
        "   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundCharts>\015\012  <"
        "x:FormatValue>\015\012   <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012   <x:"
        "Data>2</x:Data>\015\012  </x:FormatValue>\015\012  <x:PivotAggOrientation>Series"
        "</x:PivotAggOrientation>\015\012  <x:HidePivotFields/>\015\012  <x:NoGrouping/>\015"
        "\012  <x:NoFiltering/>\015\012  <x:Palette>\015\012   <x:Entry>#000000</x:Entry>"
        "\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012"
        "   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:"
        "Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>"
        "#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#00000"
        "0</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:E"
        "ntry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#8080FF</x:Entry>\015\012   <x:Entry>#802060</x:Entry>\015\012   <x:En"
        "try>#FFFFA0</x:Entry>\015\012   <x:Entry>#A0E0E0</x:Entry>\015\012   <x:Entry>#6"
        "00080</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:Entry>#008080<"
        "/x:Entry>\015\012   <x:Entry>#C0C0FF</x:Entry>\015\012   <x:Entry>#000080</x:Ent"
        "ry>\015\012   <x:Entry>#FF00FF</x:Entry>\015\012   <x:Entry>#80FFFF</x:Entry>\015"
        "\012   <x:Entry>#0080FF</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012  "
        " <x:Entry>#C0FF80</x:Entry>\015\012   <x:Entry>#FFC0FF</x:Entry>\015\012   <x:En"
        "try>#FF80FF</x:Entry>\015\012  </x:Palette>\015\012  <x:DefaultFont>Arial</x:Def"
        "aultFont>\015\012  <x:Chart>\015\012   <x:PlotArea>\015\012    <x:Graph>\015\012"
        "     <x:SubType>Clustered</x:SubType>\015\012     <x:Type>Column</x:Type>\015\012"
        "     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMap>\015\012  "
        "    <x:Name>Sum of Onhand</x:Name>\015\012      <x:Caption>\015\012       <x:Dat"
        "aSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Sum of Onhand&qu"
        "ot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index>0</x:Index>\015\012"
        "      <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012"
        "      </x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x"
        ":DataSourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012  "
        "     <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015"
        "\012      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None<"
        "/x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012"
        "      <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSou"
        "rceIndex>\015\012      <x:Identifier>!.Sum of Onhand</x:Identifier>\015\012     "
        "</x:Series>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:F"
        "ormatMap>\015\012      <x:Name>Sum of Min</x:Name>\015\012      <x:Caption>\015\012"
        "       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Sum"
        " of Min&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index>1</x:Inde"
        "x>\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceInd"
        "ex>\015\012      </x:Category>\015\012      <x:Value>\015\012       <x:DataSourc"
        "eIndex>0</x:DataSourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue"
        ">\015\012       <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>"
        "2</x:Data>\015\012      </x:FormatValue>\015\012      <x:Marker>\015\012       <"
        "x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:E"
        "xplode>\015\012      <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceInd"
        "ex>0</x:DataSourceIndex>\015\012      <x:Identifier>!.Sum of Min</x:Identifier>\015"
        "\012     </x:Series>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012  "
        "    </x:FormatMap>\015\012      <x:Name>Sum of Max</x:Name>\015\012      <x:Capt"
        "ion>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Da"
        "ta>&quot;Sum of Max&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Ind"
        "ex>2</x:Index>\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x:D"
        "ataSourceIndex>\015\012      </x:Category>\015\012      <x:Value>\015\012       "
        "<x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Value>\015\012      <x"
        ":FormatValue>\015\012       <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012   "
        "    <x:Data>2</x:Data>\015\012      </x:FormatValue>\015\012      <x:Marker>\015"
        "\012       <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012      <x:Ex"
        "plode>0</x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015\012      <x:D"
        "ataSourceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!.Sum of Max</x:"
        "Identifier>\015\012     </x:Series>\015\012     <x:Series>\015\012      <x:Forma"
        "tMap>\015\012      </x:FormatMap>\015\012      <x:Name>Sum of CurrMM</x:Name>\015"
        "\012      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015"
        "\012       <x:Data>&quot;Sum of CurrMM&quot;</x:Data>\015\012      </x:Caption>\015"
        "\012      <x:Index>3</x:Index>\015\012      <x:Category>\015\012       <x:DataSo"
        "urceIndex>0</x:DataSourceIndex>\015\012      </x:Category>\015\012      <x:Value"
        ">\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Value>"
        "\015\012      <x:FormatValue>\015\012       <x:DataSourceIndex>-3</x:DataSourceI"
        "ndex>\015\012       <x:Data>2</x:Data>\015\012      </x:FormatValue>\015\012    "
        "  <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015"
        "\012      <x:Explode>0</x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015"
        "\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!."
        "Sum of CurrMM</x:Identifier>\015\012     </x:Series>\015\012     <x:Series>\015\012"
        "      <x:FormatMap>\015\012      </x:FormatMap>\015\012      <x:Name>Sum of Jan<"
        "/x:Name>\015\012      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSou"
        "rceIndex>\015\012       <x:Data>&quot;Sum of Jan&quot;</x:Data>\015\012      </x"
        ":Caption>\015\012      <x:Index>4</x:Index>\015\012      <x:Category>\015\012   "
        "    <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Category>\015\012 "
        "     <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012   "
        "   </x:Value>\015\012      <x:FormatValue>\015\012       <x:DataSourceIndex>-3</"
        "x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012      </x:FormatValue"
        ">\015\012      <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015\012      <"
        "/x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012      <x:Thickness>10</x"
        ":Thickness>\015\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <"
        "x:Identifier>!.Sum of Jan</x:Identifier>\015\012     </x:Series>\015\012     <x:"
        "Series>\015\012      <x:FormatMap>\015\012      </x:FormatMap>\015\012      <x:N"
        "ame>Sum of Feb</x:Name>\015\012      <x:Caption>\015\012       <x:DataSourceInde"
        "x>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Sum of Feb&quot;</x:Data>\015"
        "\012      </x:Caption>\015\012      <x:Index>5</x:Index>\015\012      <x:Categor"
        "y>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Categ"
        "ory>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSourceInde"
        "x>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <x:DataSou"
        "rceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012      </"
        "x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015"
        "\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012      <x:Thic"
        "kness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015"
        "\012      <x:Identifier>!.Sum of Feb</x:Identifier>\015\012     </x:Series>\015\012"
        "     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMap>\015\012  "
        "    <x:Name>Sum of Mar</x:Name>\015\012      <x:Caption>\015\012       <x:DataSo"
        "urceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Sum of Mar&quot;</x"
        ":Data>\015\012      </x:Caption>\015\012      <x:Index>6</x:Index>\015\012      "
        "<x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012     "
        " </x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:Data"
        "SourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       "
        "<x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!.Sum of Mar</x:Identifier>\015\012     </x:Se"
        "ries>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatM"
        "ap>\015\012      <x:Name>Sum of Apr</x:Name>\015\012      <x:Caption>\015\012   "
        "    <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Sum of"
        " Apr&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index>7</x:Index>\015"
        "\012      <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015"
        "\012      </x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>"
        "0</x:DataSourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012"
        "       <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data"
        ">\015\012      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>"
        "None</x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015"
        "\012      <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:Dat"
        "aSourceIndex>\015\012      <x:Identifier>!.Sum of Apr</x:Identifier>\015\012    "
        " </x:Series>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:"
        "FormatMap>\015\012      <x:Name>Sum of May</x:Name>\015\012      <x:Caption>\015"
        "\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot"
        ";Sum of May&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index>8</x:"
        "Index>\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourc"
        "eIndex>\015\012      </x:Category>\015\012      <x:Value>\015\012       <x:DataS"
        "ourceIndex>0</x:DataSourceIndex>\015\012      </x:Value>\015\012      <x:FormatV"
        "alue>\015\012       <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:D"
        "ata>2</x:Data>\015\012      </x:FormatValue>\015\012      <x:Marker>\015\012    "
        "   <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>0<"
        "/x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015\012      <x:DataSourc"
        "eIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!.Sum of May</x:Identifi"
        "er>\015\012     </x:Series>\015\012     <x:Series>\015\012      <x:FormatMap>\015"
        "\012      </x:FormatMap>\015\012      <x:Name>Sum of Jun</x:Name>\015\012      <"
        "x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012      "
        " <x:Data>&quot;Sum of Jun&quot;</x:Data>\015\012      </x:Caption>\015\012      "
        "<x:Index>9</x:Index>\015\012      <x:Category>\015\012       <x:DataSourceIndex>"
        "0</x:DataSourceIndex>\015\012      </x:Category>\015\012      <x:Value>\015\012 "
        "      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Value>\015\012  "
        "    <x:FormatValue>\015\012       <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012"
        "       <x:Data>2</x:Data>\015\012      </x:FormatValue>\015\012      <x:Marker>\015"
        "\012       <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012      <x:Ex"
        "plode>0</x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015\012      <x:D"
        "ataSourceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!.Sum of Jun</x:"
        "Identifier>\015\012     </x:Series>\015\012     <x:Series>\015\012      <x:Forma"
        "tMap>\015\012      </x:FormatMap>\015\012      <x:Name>Sum of Jul</x:Name>\015\012"
        "      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012"
        "       <x:Data>&quot;Sum of Jul&quot;</x:Data>\015\012      </x:Caption>\015\012"
        "      <x:Index>10</x:Index>\015\012      <x:Category>\015\012       <x:DataSourc"
        "eIndex>0</x:DataSourceIndex>\015\012      </x:Category>\015\012      <x:Value>\015"
        "\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Value>\015\012"
        "      <x:FormatValue>\015\012       <x:DataSourceIndex>-3</x:DataSourceIndex>\015"
        "\012       <x:Data>2</x:Data>\015\012      </x:FormatValue>\015\012      <x:Mark"
        "er>\015\012       <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012    "
        "  <x:Explode>0</x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015\012   "
        "   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!.Sum of "
        "Jul</x:Identifier>\015\012     </x:Series>\015\012     <x:Series>\015\012      <"
        "x:FormatMap>\015\012      </x:FormatMap>\015\012      <x:Name>Sum of Aug</x:Name"
        ">\015\012      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourceInde"
        "x>\015\012       <x:Data>&quot;Sum of Aug&quot;</x:Data>\015\012      </x:Captio"
        "n>\015\012      <x:Index>11</x:Index>\015\012      <x:Category>\015\012       <x"
        ":DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Category>\015\012      <"
        "x:Value>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x"
        ":Value>\015\012      <x:FormatValue>\015\012       <x:DataSourceIndex>-3</x:Data"
        "SourceIndex>\015\012       <x:Data>2</x:Data>\015\012      </x:FormatValue>\015\012"
        "      <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015\012      </x:Marker"
        ">\015\012      <x:Explode>0</x:Explode>\015\012      <x:Thickness>10</x:Thicknes"
        "s>\015\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <x:Identif"
        "ier>!.Sum of Aug</x:Identifier>\015\012     </x:Series>\015\012     <x:Series>\015"
        "\012      <x:FormatMap>\015\012      </x:FormatMap>\015\012      <x:Name>Sum of "
        "Sep</x:Name>\015\012      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:Dat"
        "aSourceIndex>\015\012       <x:Data>&quot;Sum of Sep&quot;</x:Data>\015\012     "
        " </x:Caption>\015\012      <x:Index>12</x:Index>\015\012      <x:Category>\015\012"
        "       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Category>\015\012"
        "      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  "
        "    </x:Value>\015\012      <x:FormatValue>\015\012       <x:DataSourceIndex>-3<"
        "/x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012      </x:FormatValu"
        "e>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015\012      "
        "</x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012      <x:Thickness>10</"
        "x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "<x:Identifier>!.Sum of Sep</x:Identifier>\015\012     </x:Series>\015\012     <x"
        ":Series>\015\012      <x:FormatMap>\015\012      </x:FormatMap>\015\012      <x:"
        "Name>Sum of Oct</x:Name>\015\012      <x:Caption>\015\012       <x:DataSourceInd"
        "ex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Sum of Oct&quot;</x:Data>\015"
        "\012      </x:Caption>\015\012      <x:Index>13</x:Index>\015\012      <x:Catego"
        "ry>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Cate"
        "gory>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSourceInd"
        "ex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <x:DataSo"
        "urceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012      <"
        "/x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015"
        "\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012      <x:Thic"
        "kness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015"
        "\012      <x:Identifier>!.Sum of Oct</x:Identifier>\015\012     </x:Series>\015\012"
        "     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMap>\015\012  "
        "    <x:Name>Sum of Nov</x:Name>\015\012      <x:Caption>\015\012       <x:DataSo"
        "urceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Sum of Nov&quot;</x"
        ":Data>\015\012      </x:Caption>\015\012      <x:Index>14</x:Index>\015\012     "
        " <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012    "
        "  </x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:Dat"
        "aSourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012      "
        " <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!.Sum of Nov</x:Identifier>\015\012     </x:Se"
        "ries>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatM"
        "ap>\015\012      <x:Name>Sum of Dec</x:Name>\015\012      <x:Caption>\015\012   "
        "    <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Sum of"
        " Dec&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index>15</x:Index>"
        "\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex"
        ">\015\012      </x:Category>\015\012      <x:Value>\015\012       <x:DataSourceI"
        "ndex>0</x:DataSourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015"
        "\012       <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:"
        "Data>\015\012      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Sym"
        "bol>None</x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explod"
        "e>\015\012      <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0<"
        "/x:DataSourceIndex>\015\012      <x:Identifier>!.Sum of Dec</x:Identifier>\015\012"
        "     </x:Series>\015\012     <x:Dimension>\015\012      <x:ScaleID>256307236</x:"
        "ScaleID>\015\012      <x:Index>Categories</x:Index>\015\012     </x:Dimension>\015"
        "\012     <x:Dimension>\015\012      <x:ScaleID>256307456</x:ScaleID>\015\012    "
        "  <x:Index>Value</x:Index>\015\012     </x:Dimension>\015\012     <x:Dimension>\015"
        "\012      <x:ScaleID>256307676</x:ScaleID>\015\012      <x:Index>FormatValue</x:"
        "Index>\015\012     </x:Dimension>\015\012     <x:Overlap>0</x:Overlap>\015\012  "
        "   <x:GapWidth>150</x:GapWidth>\015\012     <x:FirstSliceAngle>0</x:FirstSliceAn"
        "gle>\015\012    </x:Graph>\015\012    <x:Axis>\015\012     <x:AxisID>256495128</"
        "x:AxisID>\015\012     <x:ScaleID>256307456</x:ScaleID>\015\012     <x:Type>Value"
        "</x:Type>\015\012     <x:MajorGridlines>\015\012     </x:MajorGridlines>\015\012"
        "     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTi"
        "ck>\015\012     <x:Title>\015\012      <x:Font>\015\012       <x:FontName>Trebuc"
        "het MS</x:FontName>\015\012       <x:Size>10</x:Size>\015\012      </x:Font>\015"
        "\012      <x:Border>\015\012       <x:ColorIndex>None</x:ColorIndex>\015\012    "
        "  </x:Border>\015\012      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:Da"
        "taSourceIndex>\015\012       <x:Data>&quot;Units Sold&quot;</x:Data>\015\012    "
        "  </x:Caption>\015\012      <x:Position>Left</x:Position>\015\012     </x:Title>"
        "\015\012     <x:Placement>Left</x:Placement>\015\012    </x:Axis>\015\012    <x:"
        "Axis>\015\012     <x:AxisID>256495884</x:AxisID>\015\012     <x:ScaleID>25630723"
        "6</x:ScaleID>\015\012     <x:Type>Category</x:Type>\015\012     <x:Number>\015\012"
        "      <x:FormatString>General</x:FormatString>\015\012     </x:Number>\015\012  "
        "   <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick"
        ">\015\012     <x:Title>\015\012      <x:Font>\015\012       <x:FontName>Trebuche"
        "t MS</x:FontName>\015\012       <x:Size>10</x:Size>\015\012      </x:Font>\015\012"
        "      <x:Border>\015\012       <x:ColorIndex>None</x:ColorIndex>\015\012      </"
        "x:Border>\015\012      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSo"
        "urceIndex>\015\012       <x:Data>&quot;Onh    Min      Max      Cur       Jan   "
        "    Feb      Mar      Apr     May      Jun      Jul      Aug      Sep      Oct  "
        "    Nov      Dec&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Positi"
        "on>Bottom</x:Position>\015\012     </x:Title>\015\012     <x:Placement>Bottom</x"
        ":Placement>\015\012     <x:GroupingEnum>Auto</x:GroupingEnum>\015\012    </x:Axi"
        "s>\015\012   </x:PlotArea>\015\012   <x:Identifier></x:Identifier>\015\012  </x:"
        "Chart>\015\012  <x:Scaling>\015\012   <x:ScaleID>256307236</x:ScaleID>\015\012  "
        "</x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>256307456</x:ScaleID>\015"
        "\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>256307676</x:ScaleI"
        "D>\015\012  </x:Scaling>\015\012  <x:HidePassiveAlert/>\015\012 </x:ChartSpace>\015"
        "\012</xml>"
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
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =60
                    Width =840
                    Height =240
                    Name ="Code_Label"
                    Caption ="Code"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =960
                    Top =60
                    Width =2160
                    Height =240
                    Name ="Description_Label"
                    Caption ="Description"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =3180
                    Top =60
                    Width =840
                    Height =240
                    Name ="CurrMM_Label"
                    Caption ="CurrMM"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =4080
                    Top =60
                    Width =2160
                    Height =240
                    Name ="Jan_Label"
                    Caption ="Jan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =6300
                    Top =60
                    Width =2160
                    Height =240
                    Name ="Feb_Label"
                    Caption ="Feb"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =8520
                    Top =60
                    Width =2160
                    Height =240
                    Name ="Mar_Label"
                    Caption ="Mar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =10740
                    Top =60
                    Width =2160
                    Height =240
                    Name ="Apr_Label"
                    Caption ="Apr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =12960
                    Top =60
                    Width =2160
                    Height =240
                    Name ="May_Label"
                    Caption ="May"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =15180
                    Top =60
                    Width =2160
                    Height =240
                    Name ="Jun_Label"
                    Caption ="Jun"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =17400
                    Top =60
                    Width =2160
                    Height =240
                    Name ="Jul_Label"
                    Caption ="Jul"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =19620
                    Top =60
                    Width =2160
                    Height =240
                    Name ="Aug_Label"
                    Caption ="Aug"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =21840
                    Top =60
                    Width =2160
                    Height =240
                    Name ="Sep_Label"
                    Caption ="Sep"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =24060
                    Top =60
                    Width =2160
                    Height =240
                    Name ="Oct_Label"
                    Caption ="Oct"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =26280
                    Top =60
                    Width =2160
                    Height =240
                    Name ="Nov_Label"
                    Caption ="Nov"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =28500
                    Top =60
                    Width =2310
                    Height =240
                    Name ="Dec_Label"
                    Caption ="Dec"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =1140
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =900
                    ColumnOrder =0
                    Name ="Code"
                    ControlSource ="Code"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =60
                    Width =1920
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7560
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =900
                    ColumnOrder =5
                    TabIndex =5
                    Name ="CurrMM"
                    ControlSource ="CurrMM"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8460
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =6
                    TabIndex =6
                    Name ="Jan"
                    ControlSource ="Jan"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9360
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =7
                    TabIndex =7
                    Name ="Feb"
                    ControlSource ="Feb"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10310
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =8
                    TabIndex =8
                    Name ="Mar"
                    ControlSource ="Mar"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12144
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =9
                    TabIndex =9
                    Name ="Apr"
                    ControlSource ="Apr"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13978
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =10
                    TabIndex =10
                    Name ="May"
                    ControlSource ="May"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15812
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =11
                    TabIndex =11
                    Name ="Jun"
                    ControlSource ="Jun"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =17646
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =12
                    TabIndex =12
                    Name ="Jul"
                    ControlSource ="Jul"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =19480
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =13
                    TabIndex =13
                    Name ="Aug"
                    ControlSource ="Aug"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =21314
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =14
                    TabIndex =14
                    Name ="Sep"
                    ControlSource ="Sep"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =23148
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =15
                    TabIndex =15
                    Name ="Oct"
                    ControlSource ="Oct"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =24982
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =16
                    TabIndex =16
                    Name ="Nov"
                    ControlSource ="Nov"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =26816
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    ColumnOrder =17
                    TabIndex =17
                    Name ="Dec"
                    ControlSource ="Dec"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =60
                    Width =780
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Onhand"
                    ControlSource ="Onhand"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4020
                            Top =420
                            Width =690
                            Height =240
                            Name ="Label30"
                            Caption ="Onhand"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4980
                    Top =60
                    Width =780
                    ColumnOrder =3
                    TabIndex =3
                    Name ="OrderAt"
                    ControlSource ="OrderAt"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5040
                            Top =360
                            Width =660
                            Height =240
                            Name ="Label31"
                            Caption ="Min"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6120
                    Top =60
                    Width =780
                    ColumnOrder =4
                    TabIndex =4
                    Name ="OrderTo"
                    ControlSource ="OrderTo"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6120
                            Top =360
                            Width =705
                            Height =240
                            Name ="Label32"
                            Caption ="Max"
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
' See "ItemChartMovement.cls"
