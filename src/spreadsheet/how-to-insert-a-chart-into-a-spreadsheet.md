# Insert a chart into a spreadsheet document

This topic shows how to use the classes in the Open XML SDK for Office to insert a chart into a spreadsheet document programmatically.

## Row element

In this how-to, you are going to deal with the row, cell, and cell value
elements. Therefore it is useful to familiarize yourself with these
elements. The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces row (`<row/>`) element.

> The row element expresses information about an entire row of a
> worksheet, and contains all cell definitions for a particular row in
> the worksheet.
>
> This row expresses information about row 2 in the worksheet, and
> contains 3 cell definitions.

```xml
    <row r="2" spans="2:12">
      <c r="C2" s="1">
        <f>PMT(B3/12,B4,-B5)</f>
        <v>672.68336574300008</v>
      </c>
      <c r="D2">
        <v>180</v>
      </c>
      <c r="E2">
        <v>360</v>
      </c>
    </row>
```

> &copy; ISO/IEC 29500: 2016

The following XML Schema code example defines the contents of the row
element.

```xml
    <complexType name="CT_Row">
       <sequence>
           <element name="c" type="CT_Cell" minOccurs="0" maxOccurs="unbounded"/>
           <element name="extLst" minOccurs="0" type="CT_ExtensionList"/>
       </sequence>
       <attribute name="r" type="xsd:unsignedInt" use="optional"/>
       <attribute name="spans" type="ST_CellSpans" use="optional"/>
       <attribute name="s" type="xsd:unsignedInt" use="optional" default="0"/>
       <attribute name="customFormat" type="xsd:boolean" use="optional" default="false"/>
       <attribute name="ht" type="xsd:double" use="optional"/>
       <attribute name="hidden" type="xsd:boolean" use="optional" default="false"/>
       <attribute name="customHeight" type="xsd:boolean" use="optional" default="false"/>
       <attribute name="outlineLevel" type="xsd:unsignedByte" use="optional" default="0"/>
       <attribute name="collapsed" type="xsd:boolean" use="optional" default="false"/>
       <attribute name="thickTop" type="xsd:boolean" use="optional" default="false"/>
       <attribute name="thickBot" type="xsd:boolean" use="optional" default="false"/>
       <attribute name="ph" type="xsd:boolean" use="optional" default="false"/>
    </complexType>
```

## Cell element

The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces cell (`<c/>`) element.

> This collection represents a cell in the worksheet. Information about
> the cell's location (reference), value, data type, formatting, and
> formula is expressed here.
>
> This example shows the information stored for a cell whose address in
> the grid is C6, whose style index is 6, and whose value metadata index
> is 15. The cell contains a formula as well as a calculated result of
> that formula.

```xml
    <c r="C6" s="1" vm="15">
      <f>CUBEVALUE("xlextdat9 Adventure Works",C$5,$A6)</f>
      <v>2838512.355</v>
    </c>
```

> &copy; ISO/IEC 29500: 2016

The following XML Schema code example defines the contents of this
element.

```xml
    <complexType name="CT_Cell">
       <sequence>
           <element name="f" type="CT_CellFormula" minOccurs="0" maxOccurs="1"/>
           <element name="v" type="ST_Xstring" minOccurs="0" maxOccurs="1"/>
           <element name="is" type="CT_Rst" minOccurs="0" maxOccurs="1"/>
           <element name="extLst" minOccurs="0" type="CT_ExtensionList"/>
       </sequence>
       <attribute name="r" type="ST_CellRef" use="optional"/>
       <attribute name="s" type="xsd:unsignedInt" use="optional" default="0"/>
       <attribute name="t" type="ST_CellType" use="optional" default="n"/>
       <attribute name="cm" type="xsd:unsignedInt" use="optional" default="0"/>
       <attribute name="vm" type="xsd:unsignedInt" use="optional" default="0"/>
       <attribute name="ph" type="xsd:boolean" use="optional" default="false"/>
    </complexType>
```

## Cell value element

The following text from the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification
introduces Cell Value (`<c/>`) element.

> This element expresses the value contained in a cell. If the cell
> contains a string, then this value is an index into the shared string
> table, pointing to the actual string value. Otherwise, the value of
> the cell is expressed directly in this element. Cells containing
> formulas express the last calculated result of the formula in this
> element.
>
> For applications not wanting to implement the shared string table, an
> "inline string" may be expressed in an `<is/>` element under `<c/>` (instead of a `<v/>` element under `<c/>`), in the same way a string would be
> expressed in the shared string table.
>
> &copy; ISO/IEC 29500: 2016

In the following example cell B4 contains the number 360.

```xml
    <c r="B4">
      <v>360</v>
    </c>
```

## How the sample code works

After opening the spreadsheet file for read/write access, the code verifies if the specified worksheet exists. It then adds a new `DocumentFormat.OpenXml.Packaging.DrawingsPart` object using the `DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer.AddNewPart` method, appends it to the worksheet, and saves the worksheet part. The code then adds a new `DocumentFormat.OpenXml.Packaging.ChartPart` object, appends a new `DocumentFormat.OpenXml.Packaging.ChartPart.ChartSpace` object to the `ChartPart` object, and then appends a new `DocumentFormat.OpenXml.Drawing.Charts.ChartSpace.EditingLanguage` object to the `ChartSpace` object that specifies the language for the chart is English-US.

### [C#](#tab/cs-1)
```csharp
        IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);

        if (sheets is null || sheets.Count() == 0)
        {
            // The specified worksheet does not exist.
            return;
        }

        string? id = sheets.First().Id;

        if (id is null)
        {
            // The worksheet does not have an ID.
            return;
        }

        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(id);

        // Add a new drawing to the worksheet.
        DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
        worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing()
        { Id = worksheetPart.GetIdOfPart(drawingsPart) });

        // Add a new chart and set the chart language to English-US.
        ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = new ChartSpace();
        chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
        DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
            new DocumentFormat.OpenXml.Drawing.Charts.Chart());
```

### [Visual Basic](#tab/vb-1)
```vb
            Dim sheets As IEnumerable(Of Sheet) = document.WorkbookPart?.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = worksheetName)

            If sheets Is Nothing OrElse sheets.Count() = 0 Then
                ' The specified worksheet does not exist.
                Return
            End If

            Dim id As String = sheets.First().Id

            If id Is Nothing Then
                ' The worksheet does not have an ID.
                Return
            End If

            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(id), WorksheetPart)

            ' Add a new drawing to the worksheet.
            Dim drawingsPart As DrawingsPart = worksheetPart.AddNewPart(Of DrawingsPart)()
            worksheetPart.Worksheet.Append(New DocumentFormat.OpenXml.Spreadsheet.Drawing() With {.Id = worksheetPart.GetIdOfPart(drawingsPart)})

            ' Add a new chart and set the chart language to English-US.
            Dim chartPart As ChartPart = drawingsPart.AddNewPart(Of ChartPart)()
            chartPart.ChartSpace = New ChartSpace()
            chartPart.ChartSpace.Append(New EditingLanguage() With {.Val = New StringValue("en-US")})
            Dim chart As DocumentFormat.OpenXml.Drawing.Charts.Chart = chartPart.ChartSpace.AppendChild(Of DocumentFormat.OpenXml.Drawing.Charts.Chart)(New DocumentFormat.OpenXml.Drawing.Charts.Chart())
```
***

The code creates a new clustered column chart by creating a new `DocumentFormat.OpenXml.Drawing.Charts.BarChart` object with
`DocumentFormat.OpenXml.Drawing.Charts.BarDirectionValues` object set to `Column` and `DocumentFormat.OpenXml.Drawing.Charts.BarGroupingValues` object set to `Clustered`.

The code then iterates through each key in the `Dictionary` class. For each key, it appends a
`DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries` object to the `BarChart` object and sets the `DocumentFormat.OpenXml.Drawing.Charts.SeriesText` object of the `BarChartSeries` object to equal the key. For each key, it appends a `DocumentFormat.OpenXml.Drawing.Charts.NumberLiteral` object to the `Values` collection of the `BarChartSeries` object and sets the `NumberLiteral` object to equal the `Dictionary` class value corresponding to the key.

### [C#](#tab/cs-2)
```csharp
        // Create a new clustered column chart.
        PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
        Layout layout = plotArea.AppendChild<Layout>(new Layout());
        BarChart barChart = plotArea.AppendChild<BarChart>(new BarChart(new BarDirection()
        { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
            new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) }));

        uint i = 0;

        // Iterate through each key in the Dictionary collection and add the key to the chart Series
        // and add the corresponding value to the chart Values.
        foreach (string key in data.Keys)
        {
            BarChartSeries barChartSeries = barChart.AppendChild<BarChartSeries>(new BarChartSeries(new Index()
            {
                Val = new UInt32Value(i)
            },
                new Order() { Val = new UInt32Value(i) },
                new SeriesText(new NumericValue() { Text = key })));

            StringLiteral strLit = barChartSeries.AppendChild<CategoryAxisData>(new CategoryAxisData()).AppendChild<StringLiteral>(new StringLiteral());
            strLit.Append(new PointCount() { Val = new UInt32Value(1U) });
            strLit.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(0U) }).Append(new NumericValue(title));

            NumberLiteral numLit = barChartSeries.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(
                new DocumentFormat.OpenXml.Drawing.Charts.Values()).AppendChild<NumberLiteral>(new NumberLiteral());
            numLit.Append(new FormatCode("General"));
            numLit.Append(new PointCount() { Val = new UInt32Value(1U) });
            numLit.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(0u) }).Append
(new NumericValue(data[key].ToString()));

            i++;
        }
```

### [Visual Basic](#tab/vb-2)
```vb
            ' Create a new clustered column chart.
            Dim plotArea As PlotArea = chart.AppendChild(Of PlotArea)(New PlotArea())
            Dim layout As Layout = plotArea.AppendChild(Of Layout)(New Layout())
            Dim barChart As BarChart = plotArea.AppendChild(Of BarChart)(New BarChart(New BarDirection() With {.Val = New EnumValue(Of BarDirectionValues)(BarDirectionValues.Column)}, New BarGrouping() With {.Val = New EnumValue(Of BarGroupingValues)(BarGroupingValues.Clustered)}))

            Dim i As UInteger = 0

            ' Iterate through each key in the Dictionary collection and add the key to the chart Series
            ' and add the corresponding value to the chart Values.
            For Each key As String In data.Keys
                Dim barChartSeries As BarChartSeries = barChart.AppendChild(Of BarChartSeries)(New BarChartSeries(New Index() With {.Val = New UInt32Value(i)}, New Order() With {.Val = New UInt32Value(i)}, New SeriesText(New NumericValue() With {.Text = key})))

                Dim strLit As StringLiteral = barChartSeries.AppendChild(Of CategoryAxisData)(New CategoryAxisData()).AppendChild(Of StringLiteral)(New StringLiteral())
                strLit.Append(New PointCount() With {.Val = New UInt32Value(1UI)})
                strLit.AppendChild(Of StringPoint)(New StringPoint() With {.Index = New UInt32Value(0UI)}).Append(New NumericValue(title))

                Dim numLit As NumberLiteral = barChartSeries.AppendChild(Of DocumentFormat.OpenXml.Drawing.Charts.Values)(New DocumentFormat.OpenXml.Drawing.Charts.Values()).AppendChild(Of NumberLiteral)(New NumberLiteral())
                numLit.Append(New FormatCode("General"))
                numLit.Append(New PointCount() With {.Val = New UInt32Value(1UI)})
                numLit.AppendChild(Of NumericPoint)(New NumericPoint() With {.Index = New UInt32Value(0UI)}).Append(New NumericValue(data(key).ToString()))

                i += 1
            Next
```
***

The code adds the `DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis` object and `DocumentFormat.OpenXml.Drawing.Charts.ValueAxis` object to the chart and sets the value of the following properties: `DocumentFormat.OpenXml.Drawing.Charts.Scaling`, `DocumentFormat.OpenXml.Drawing.Charts.AxisPosition`, `DocumentFormat.OpenXml.Drawing.Charts.TickLabelPosition`, `DocumentFormat.OpenXml.Drawing.Charts.CrossingAxis`, `DocumentFormat.OpenXml.Drawing.Charts.Crosses`, `DocumentFormat.OpenXml.Drawing.Charts.AutoLabeled`, `DocumentFormat.OpenXml.Drawing.Charts.LabelAlignment`, and `DocumentFormat.OpenXml.Drawing.Charts.LabelOffset`. It also adds the `DocumentFormat.OpenXml.Drawing.Charts.Chart.Legend` object to the chart and saves the chart part.

### [C#](#tab/cs-3)
```csharp
        barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
        barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

        // Add the Category Axis.
        CategoryAxis catAx = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new AxisId()
        { Val = new UInt32Value(48650112u) }, new Scaling(new Orientation()
        {
            Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
        }),
            new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
            new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
            new CrossingAxis() { Val = new UInt32Value(48672768U) },
            new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
            new AutoLabeled() { Val = new BooleanValue(true) },
            new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
            new LabelOffset() { Val = new UInt16Value((ushort)100) }));

        // Add the Value Axis.
        ValueAxis valAx = plotArea.AppendChild<ValueAxis>(new ValueAxis(new AxisId() { Val = new UInt32Value(48672768u) },
            new Scaling(new Orientation()
            {
                Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
            }),
            new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
            new MajorGridlines(),
            new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
            {
                FormatCode = new StringValue("General"),
                SourceLinked = new BooleanValue(true)
            }, new TickLabelPosition()
            {
                Val = new EnumValue<TickLabelPositionValues>
(TickLabelPositionValues.NextTo)
            }, new CrossingAxis() { Val = new UInt32Value(48650112U) },
            new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
            new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));

        // Add the chart Legend.
        Legend legend = chart.AppendChild<Legend>(new Legend(new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Right) },
            new Layout()));

        chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });
```

### [Visual Basic](#tab/vb-3)
```vb
            barChart.Append(New AxisId() With {.Val = New UInt32Value(48650112UI)})
            barChart.Append(New AxisId() With {.Val = New UInt32Value(48672768UI)})

            ' Add the Category Axis.
            Dim catAx As CategoryAxis = plotArea.AppendChild(Of CategoryAxis)(New CategoryAxis(New AxisId() With {.Val = New UInt32Value(48650112UI)}, New Scaling(New Orientation() With {.Val = New EnumValue(Of DocumentFormat.OpenXml.Drawing.Charts.OrientationValues)(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)}), New AxisPosition() With {.Val = New EnumValue(Of AxisPositionValues)(AxisPositionValues.Bottom)}, New TickLabelPosition() With {.Val = New EnumValue(Of TickLabelPositionValues)(TickLabelPositionValues.NextTo)}, New CrossingAxis() With {.Val = New UInt32Value(48672768UI)}, New Crosses() With {.Val = New EnumValue(Of CrossesValues)(CrossesValues.AutoZero)}, New AutoLabeled() With {.Val = New BooleanValue(True)}, New LabelAlignment() With {.Val = New EnumValue(Of LabelAlignmentValues)(LabelAlignmentValues.Center)}, New LabelOffset() With {.Val = New UInt16Value(CUShort(100))}))

            ' Add the Value Axis.
            Dim valAx As ValueAxis = plotArea.AppendChild(Of ValueAxis)(New ValueAxis(New AxisId() With {.Val = New UInt32Value(48672768UI)}, New Scaling(New Orientation() With {.Val = New EnumValue(Of DocumentFormat.OpenXml.Drawing.Charts.OrientationValues)(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)}), New AxisPosition() With {.Val = New EnumValue(Of AxisPositionValues)(AxisPositionValues.Left)}, New MajorGridlines(), New DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() With {.FormatCode = New StringValue("General"), .SourceLinked = New BooleanValue(True)}, New TickLabelPosition() With {.Val = New EnumValue(Of TickLabelPositionValues)(TickLabelPositionValues.NextTo)}, New CrossingAxis() With {.Val = New UInt32Value(48650112UI)}, New Crosses() With {.Val = New EnumValue(Of CrossesValues)(CrossesValues.AutoZero)}, New CrossBetween() With {.Val = New EnumValue(Of CrossBetweenValues)(CrossBetweenValues.Between)}))

            ' Add the chart Legend.
            Dim legend As Legend = chart.AppendChild(Of Legend)(New Legend(New LegendPosition() With {.Val = New EnumValue(Of LegendPositionValues)(LegendPositionValues.Right)}, New Layout()))

            chart.Append(New PlotVisibleOnly() With {.Val = New BooleanValue(True)})
```
***

The code positions the chart on the worksheet by creating a `DocumentFormat.OpenXml.Packaging.DrawingsPart.WorksheetDrawing` object and appending a `TwoCellAnchor` object. The `TwoCellAnchor` object specifies how to move or resize the chart if you move the rows and columns between the `DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker` and `DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker` anchors. The code then creates a `DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame` object to contain the chart and names the chart "Chart 1".

### [C#](#tab/cs-4)
```csharp
        // Position the chart on the worksheet using a TwoCellAnchor object.
        drawingsPart.WorksheetDrawing = new WorksheetDrawing();
        TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());
        twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId("9"),
            new ColumnOffset("581025"),
            new RowId("17"),
            new RowOffset("114300")));
        twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId("17"),
            new ColumnOffset("276225"),
            new RowId("32"),
            new RowOffset("0")));

        // Append a GraphicFrame to the TwoCellAnchor object.
        DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame =
            twoCellAnchor.AppendChild<DocumentFormat.OpenXml.
Drawing.Spreadsheet.GraphicFrame>(new DocumentFormat.OpenXml.Drawing.
Spreadsheet.GraphicFrame());
        graphicFrame.Macro = "";

        graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = new UInt32Value(2u), Name = "Chart 1" },
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

        graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L },
                                                                new Extents() { Cx = 0L, Cy = 0L }));

        graphicFrame.Append(new Graphic(new GraphicData(new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) })
        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

        twoCellAnchor.Append(new ClientData());
```

### [Visual Basic](#tab/vb-4)
```vb
            ' Position the chart on the worksheet using a TwoCellAnchor object.
            drawingsPart.WorksheetDrawing = New WorksheetDrawing()
            Dim twoCellAnchor As TwoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(Of TwoCellAnchor)(New TwoCellAnchor())
            twoCellAnchor.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(New ColumnId("9"), New ColumnOffset("581025"), New RowId("17"), New RowOffset("114300")))
            twoCellAnchor.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(New ColumnId("17"), New ColumnOffset("276225"), New RowId("32"), New RowOffset("0")))

            ' Append a GraphicFrame to the TwoCellAnchor object.
            Dim graphicFrame As DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame = twoCellAnchor.AppendChild(Of DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame)(New DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame())
            graphicFrame.Macro = ""

            graphicFrame.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(New DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() With {.Id = New UInt32Value(2UI), .Name = "Chart 1"}, New DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()))

            graphicFrame.Append(New Transform(New Offset() With {.X = 0L, .Y = 0L}, New Extents() With {.Cx = 0L, .Cy = 0L}))

            graphicFrame.Append(New Graphic(New GraphicData(New ChartReference() With {.Id = drawingsPart.GetIdOfPart(chartPart)}) With {.Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"}))

            twoCellAnchor.Append(New ClientData())
```
***

## Sample Code

> **Note**
> This code can be run only once. You cannot create more than one instance of the chart.

The following is the complete sample code in both C\# and Visual Basic.

### [C#](#tab/cs)
```csharp
static void InsertChartInSpreadsheet(string docName, string worksheetName, string title, Dictionary<string, int> data)
{
    // Open the document for editing.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
    {
        IEnumerable<Sheet>? sheets = document.WorkbookPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);

        if (sheets is null || sheets.Count() == 0)
        {
            // The specified worksheet does not exist.
            return;
        }

        string? id = sheets.First().Id;

        if (id is null)
        {
            // The worksheet does not have an ID.
            return;
        }

        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(id);

        // Add a new drawing to the worksheet.
        DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
        worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing()
        { Id = worksheetPart.GetIdOfPart(drawingsPart) });

        // Add a new chart and set the chart language to English-US.
        ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = new ChartSpace();
        chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
        DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
            new DocumentFormat.OpenXml.Drawing.Charts.Chart());
        // Create a new clustered column chart.
        PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
        Layout layout = plotArea.AppendChild<Layout>(new Layout());
        BarChart barChart = plotArea.AppendChild<BarChart>(new BarChart(new BarDirection()
        { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
            new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) }));

        uint i = 0;

        // Iterate through each key in the Dictionary collection and add the key to the chart Series
        // and add the corresponding value to the chart Values.
        foreach (string key in data.Keys)
        {
            BarChartSeries barChartSeries = barChart.AppendChild<BarChartSeries>(new BarChartSeries(new Index()
            {
                Val = new UInt32Value(i)
            },
                new Order() { Val = new UInt32Value(i) },
                new SeriesText(new NumericValue() { Text = key })));

            StringLiteral strLit = barChartSeries.AppendChild<CategoryAxisData>(new CategoryAxisData()).AppendChild<StringLiteral>(new StringLiteral());
            strLit.Append(new PointCount() { Val = new UInt32Value(1U) });
            strLit.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(0U) }).Append(new NumericValue(title));

            NumberLiteral numLit = barChartSeries.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(
                new DocumentFormat.OpenXml.Drawing.Charts.Values()).AppendChild<NumberLiteral>(new NumberLiteral());
            numLit.Append(new FormatCode("General"));
            numLit.Append(new PointCount() { Val = new UInt32Value(1U) });
            numLit.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(0u) }).Append
(new NumericValue(data[key].ToString()));

            i++;
        }
        barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
        barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

        // Add the Category Axis.
        CategoryAxis catAx = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new AxisId()
        { Val = new UInt32Value(48650112u) }, new Scaling(new Orientation()
        {
            Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
        }),
            new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
            new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
            new CrossingAxis() { Val = new UInt32Value(48672768U) },
            new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
            new AutoLabeled() { Val = new BooleanValue(true) },
            new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
            new LabelOffset() { Val = new UInt16Value((ushort)100) }));

        // Add the Value Axis.
        ValueAxis valAx = plotArea.AppendChild<ValueAxis>(new ValueAxis(new AxisId() { Val = new UInt32Value(48672768u) },
            new Scaling(new Orientation()
            {
                Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
            }),
            new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
            new MajorGridlines(),
            new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
            {
                FormatCode = new StringValue("General"),
                SourceLinked = new BooleanValue(true)
            }, new TickLabelPosition()
            {
                Val = new EnumValue<TickLabelPositionValues>
(TickLabelPositionValues.NextTo)
            }, new CrossingAxis() { Val = new UInt32Value(48650112U) },
            new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
            new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));

        // Add the chart Legend.
        Legend legend = chart.AppendChild<Legend>(new Legend(new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Right) },
            new Layout()));

        chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });
        // Position the chart on the worksheet using a TwoCellAnchor object.
        drawingsPart.WorksheetDrawing = new WorksheetDrawing();
        TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());
        twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId("9"),
            new ColumnOffset("581025"),
            new RowId("17"),
            new RowOffset("114300")));
        twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId("17"),
            new ColumnOffset("276225"),
            new RowId("32"),
            new RowOffset("0")));

        // Append a GraphicFrame to the TwoCellAnchor object.
        DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame =
            twoCellAnchor.AppendChild<DocumentFormat.OpenXml.
Drawing.Spreadsheet.GraphicFrame>(new DocumentFormat.OpenXml.Drawing.
Spreadsheet.GraphicFrame());
        graphicFrame.Macro = "";

        graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = new UInt32Value(2u), Name = "Chart 1" },
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

        graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L },
                                                                new Extents() { Cx = 0L, Cy = 0L }));

        graphicFrame.Append(new Graphic(new GraphicData(new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) })
        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

        twoCellAnchor.Append(new ClientData());
    }

}
```

### [Visual Basic](#tab/vb)
```vb
    Sub InsertChartInSpreadsheet(docName As String, worksheetName As String, title As String, data As Dictionary(Of String, Integer))
        ' Open the document for editing.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            Dim sheets As IEnumerable(Of Sheet) = document.WorkbookPart?.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = worksheetName)

            If sheets Is Nothing OrElse sheets.Count() = 0 Then
                ' The specified worksheet does not exist.
                Return
            End If

            Dim id As String = sheets.First().Id

            If id Is Nothing Then
                ' The worksheet does not have an ID.
                Return
            End If

            Dim worksheetPart As WorksheetPart = CType(document.WorkbookPart.GetPartById(id), WorksheetPart)

            ' Add a new drawing to the worksheet.
            Dim drawingsPart As DrawingsPart = worksheetPart.AddNewPart(Of DrawingsPart)()
            worksheetPart.Worksheet.Append(New DocumentFormat.OpenXml.Spreadsheet.Drawing() With {.Id = worksheetPart.GetIdOfPart(drawingsPart)})

            ' Add a new chart and set the chart language to English-US.
            Dim chartPart As ChartPart = drawingsPart.AddNewPart(Of ChartPart)()
            chartPart.ChartSpace = New ChartSpace()
            chartPart.ChartSpace.Append(New EditingLanguage() With {.Val = New StringValue("en-US")})
            Dim chart As DocumentFormat.OpenXml.Drawing.Charts.Chart = chartPart.ChartSpace.AppendChild(Of DocumentFormat.OpenXml.Drawing.Charts.Chart)(New DocumentFormat.OpenXml.Drawing.Charts.Chart())
            ' Create a new clustered column chart.
            Dim plotArea As PlotArea = chart.AppendChild(Of PlotArea)(New PlotArea())
            Dim layout As Layout = plotArea.AppendChild(Of Layout)(New Layout())
            Dim barChart As BarChart = plotArea.AppendChild(Of BarChart)(New BarChart(New BarDirection() With {.Val = New EnumValue(Of BarDirectionValues)(BarDirectionValues.Column)}, New BarGrouping() With {.Val = New EnumValue(Of BarGroupingValues)(BarGroupingValues.Clustered)}))

            Dim i As UInteger = 0

            ' Iterate through each key in the Dictionary collection and add the key to the chart Series
            ' and add the corresponding value to the chart Values.
            For Each key As String In data.Keys
                Dim barChartSeries As BarChartSeries = barChart.AppendChild(Of BarChartSeries)(New BarChartSeries(New Index() With {.Val = New UInt32Value(i)}, New Order() With {.Val = New UInt32Value(i)}, New SeriesText(New NumericValue() With {.Text = key})))

                Dim strLit As StringLiteral = barChartSeries.AppendChild(Of CategoryAxisData)(New CategoryAxisData()).AppendChild(Of StringLiteral)(New StringLiteral())
                strLit.Append(New PointCount() With {.Val = New UInt32Value(1UI)})
                strLit.AppendChild(Of StringPoint)(New StringPoint() With {.Index = New UInt32Value(0UI)}).Append(New NumericValue(title))

                Dim numLit As NumberLiteral = barChartSeries.AppendChild(Of DocumentFormat.OpenXml.Drawing.Charts.Values)(New DocumentFormat.OpenXml.Drawing.Charts.Values()).AppendChild(Of NumberLiteral)(New NumberLiteral())
                numLit.Append(New FormatCode("General"))
                numLit.Append(New PointCount() With {.Val = New UInt32Value(1UI)})
                numLit.AppendChild(Of NumericPoint)(New NumericPoint() With {.Index = New UInt32Value(0UI)}).Append(New NumericValue(data(key).ToString()))

                i += 1
            Next
            barChart.Append(New AxisId() With {.Val = New UInt32Value(48650112UI)})
            barChart.Append(New AxisId() With {.Val = New UInt32Value(48672768UI)})

            ' Add the Category Axis.
            Dim catAx As CategoryAxis = plotArea.AppendChild(Of CategoryAxis)(New CategoryAxis(New AxisId() With {.Val = New UInt32Value(48650112UI)}, New Scaling(New Orientation() With {.Val = New EnumValue(Of DocumentFormat.OpenXml.Drawing.Charts.OrientationValues)(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)}), New AxisPosition() With {.Val = New EnumValue(Of AxisPositionValues)(AxisPositionValues.Bottom)}, New TickLabelPosition() With {.Val = New EnumValue(Of TickLabelPositionValues)(TickLabelPositionValues.NextTo)}, New CrossingAxis() With {.Val = New UInt32Value(48672768UI)}, New Crosses() With {.Val = New EnumValue(Of CrossesValues)(CrossesValues.AutoZero)}, New AutoLabeled() With {.Val = New BooleanValue(True)}, New LabelAlignment() With {.Val = New EnumValue(Of LabelAlignmentValues)(LabelAlignmentValues.Center)}, New LabelOffset() With {.Val = New UInt16Value(CUShort(100))}))

            ' Add the Value Axis.
            Dim valAx As ValueAxis = plotArea.AppendChild(Of ValueAxis)(New ValueAxis(New AxisId() With {.Val = New UInt32Value(48672768UI)}, New Scaling(New Orientation() With {.Val = New EnumValue(Of DocumentFormat.OpenXml.Drawing.Charts.OrientationValues)(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)}), New AxisPosition() With {.Val = New EnumValue(Of AxisPositionValues)(AxisPositionValues.Left)}, New MajorGridlines(), New DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() With {.FormatCode = New StringValue("General"), .SourceLinked = New BooleanValue(True)}, New TickLabelPosition() With {.Val = New EnumValue(Of TickLabelPositionValues)(TickLabelPositionValues.NextTo)}, New CrossingAxis() With {.Val = New UInt32Value(48650112UI)}, New Crosses() With {.Val = New EnumValue(Of CrossesValues)(CrossesValues.AutoZero)}, New CrossBetween() With {.Val = New EnumValue(Of CrossBetweenValues)(CrossBetweenValues.Between)}))

            ' Add the chart Legend.
            Dim legend As Legend = chart.AppendChild(Of Legend)(New Legend(New LegendPosition() With {.Val = New EnumValue(Of LegendPositionValues)(LegendPositionValues.Right)}, New Layout()))

            chart.Append(New PlotVisibleOnly() With {.Val = New BooleanValue(True)})
            ' Position the chart on the worksheet using a TwoCellAnchor object.
            drawingsPart.WorksheetDrawing = New WorksheetDrawing()
            Dim twoCellAnchor As TwoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(Of TwoCellAnchor)(New TwoCellAnchor())
            twoCellAnchor.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(New ColumnId("9"), New ColumnOffset("581025"), New RowId("17"), New RowOffset("114300")))
            twoCellAnchor.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(New ColumnId("17"), New ColumnOffset("276225"), New RowId("32"), New RowOffset("0")))

            ' Append a GraphicFrame to the TwoCellAnchor object.
            Dim graphicFrame As DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame = twoCellAnchor.AppendChild(Of DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame)(New DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame())
            graphicFrame.Macro = ""

            graphicFrame.Append(New DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(New DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() With {.Id = New UInt32Value(2UI), .Name = "Chart 1"}, New DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()))

            graphicFrame.Append(New Transform(New Offset() With {.X = 0L, .Y = 0L}, New Extents() With {.Cx = 0L, .Cy = 0L}))

            graphicFrame.Append(New Graphic(New GraphicData(New ChartReference() With {.Id = drawingsPart.GetIdOfPart(chartPart)}) With {.Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"}))

            twoCellAnchor.Append(New ClientData())
        End Using
    End Sub
```
***

## See also

[Open XML SDK class library reference](https://learn.microsoft.com/office/open-xml/open-xml-sdk)
