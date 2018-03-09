_[Workshop home](../../index.md)_  >  _[New Excel JavaScript APIs](../index.md)_ > _[Chart API additions](index.md)_ > _Introduction_



# Chart Customization

The Excel JavaScript Library provides APIs to enable your add-in to customize a chart in worksheet. To understand the concepts and the terminology of chart, please see the following articles about how users customize chart through the Excel UI:

- [Create charts](https://support.office.com/en-us/article/231c42d2-5e58-40e1-99f0-cbe618cfee1d)
- [Format charts](https://support.office.com/en-us/article/92693043-1772-46a9-90e3-88c8c76084d8)
- [Add trendlines](https://support.office.com/en-us/article/6b72b363-aa05-4c93-8c5b-22c480eb6e1f)

## Programmatic add/delete/customize series in chart

### Add series
The `ChartSeriesCollection.add` method, which will return a new created [ChartSeries](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/chartseries.md) object, is the entry point for programmatic add a series into a chart. There are two parameters to customize the new added series:

- `name` &#8212; Name of the series.
- `index` &#8212; Index value of the series to be added. Zero-based.

### Delete series

The `ChartSeries.delete` method, which will delete itslef from chart.

### Customize series
The `ChartSeries` object have below new properties and methods for user to cusomize it

| Property	   | Type	|Description| 
|:---------------|:--------|:----------|
|chartType|string|Represents the chart type of a series. Possible values are: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc..|
|doughnutHoleSize|double|Represents the doughnut hole size of a chart series.  Only valid on doughnut and doughnutExploded charts.|
|filtered|bool|Boolean value representing if the series is filtered or not. Not applicable for surface charts.|
|gapWidth|double|Represents the gap width of a chart series.  Only valid on bar and column charts, as well as|
|hasDataLabels|bool|Boolean value representing if the series has data labels or not.|
|markerBackgroundColor|string|Represents markers background color of a chart series.|
|markerForegroundColor|string|Represents markers foreground color of a chart series.|
|markerSize|int|Represents marker size of a chart series.|
|markerStyle|string|Represents marker style of a chart series. Possible values are: Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|
|plotOrder|int|Represents the plot order of a chart series within the chart group.|
|showShadow|bool|Boolean value representing if the series has shadow or not.|
|smooth|bool|Boolean value representing if the series is smooth or not. Only for line and scatter charts.|

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|setBubbleSizes(sourceData: Range)|void|Set bubble sizes for a chart series. Only works for bubble charts.|
|setValues(sourceData: Range)|void|Set values for a chart series. For scatter chart, it means Y axis values.|
|setXAxisValues(sourceData: Range)|void|Set values of X axis for a chart series. Only works for scatter charts.|


## Programmatic add/delete/customize trendline in chart

### Add trendline
The `ChartTrendlineCollection.add` method, which will return a new created [ChartTrendline](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/charttrendline.md) object, is the entry point for programmatic add a trendline into a chart. There is only one parameter to customize the new added trendline:

- `type` &#8212; Specifies the trendline type. The default value is "Linear".

### Delete trendline

The `ChartTrendline.delete` method, which will delete itslef from a ChartTrendlineCollection.

### Get a trendline

The `ChartTrendlineCollection.getItem` method, which will return a exsited trendline for a series.

- `index` &#8212; Represents the insertion order in items array. Zero-based.

### Customize trendline

The `ChartTrendline` object have below new properties for user to cusomize it

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|backward|double|Represents the number of periods that the trendline extends backward.|
|displayEquation|bool|True if the equation for the trendline is displayed on the chart.|
|displayRSquared|bool|True if the R-squared for the trendline is displayed on the chart.|
|forward|double|Represents the number of periods that the trendline extends forward.|
|intercept|object|Represents the intercept value of the trendline. Can be set to a numeric value or an empty string (for automatic values). The returned value is always a number.|
|movingAveragePeriod|int|Represents the period of a chart trendline, only for trendline with MovingAverage type.|
|name|string|Represents the name of the trendline. Can be set to a string value, or can be set to null value represents automatic values. The returned value is always a string|
|polynomialOrder|int|Represents the order of a chart trendline, only for trendline with Polynomial type.|
|type|string|Represents the type of a chart trendline. Possible values are: Linear, Expontential, Logarithmic, MovingAvg, Polynomial, Power.|


## Programmatic customize axis in chart

The `ChartAxis` object have below new properties and methods for user to customize it

### Customize axis options

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|axisBetweenCategories|bool|Represents whether value axis crosses the category axis between categories.|
|axisGroup|string|Represents the group for the specified axis. Read-only. Possible values are: Primary, Secondary.|
|baseTimeUnit|string|Returns or sets the base unit for the specified category axis. Possible values are: Days, Months, Years.|
|categoryType|string|Returns or sets the category axis type. Possible values are: Automatic, TextAxis, DateAxis.|
|crosses|string|Represents the specified axis where the other axis crosses.|
|crossesAt|double|Represents the specified axis where the other axis crosses at. Read Only. Set to this property should use SetCrossesAt(double) method. Read-only.|
|customDisplayUnit|double|Represents the custom axis display unit value. Read Only. To set this property, please use the SetCustomDisplayUnit(double) method. Read-only.|
|displayUnit|string|Represents the axis display unit. Possible values are: None, Hundreds, Thousands, TenThousands, HundredThousands, Millions, TenMillions, HundredMillons, Billions, Trillions, Custom.|
|height|double|Represents the height, in points, of the chart axis. Null if the axis's not visible. Read-only.|
|left|double|Represents the distance, in points, from the left edge of the axis to the left of chart area. Null if the axis's not visible. Read-only.|
|logBase|double|Represents the base of the logarithm when using logarithmic scales.|
|majorTimeUnitScale|string|Returns or sets the major unit scale value for the category axis when the CategoryType property is set to TimeScale. Possible values are: Days, Months, Years.|
|minorTimeUnitScale|string|Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to TimeScale. Possible values are: Days, Months, Years.|
|reversePlotOrder|bool|Represents whether Microsoft Excel plots data points from last to first.|
|scaleType|string|Represents the value axis scale type. Possible values are: Linear, Logarithmic.|
|showDisplayUnitLabel|bool|Represents whether the axis display unit label is visible.|
|top|double|Represents the distance, in points, from the top edge of the axis to the top of chart area. Null if the axis's not visible. Read-only.|
|type|string|Represents the axis type. Read-only. Possible values are: Invalid, Category, Value, SeriesAxis.|
|visible|bool|A boolean value represents the visibility of the axis.|
|width|double|Represents the width, in points, of the chart axis. Null if the axis's not visible. Read-only.|

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|setCategoryNames(sourceData: Range)|void|Sets all the category names for the specified axis.|
|setCrossesAt(value: double)|void|Set the specified axis where the other axis crosses at.|
|setCustomDisplayUnit(value: double)|void|Sets the axis display unit to a custom value.|

### Customize axis tick marks

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|majorTickMark|string|Represents the type of major tick mark for the specified axis.|
|minorTickMark|string|Represents the type of minor tick mark for the specified axis.|
|tickMarkSpacing|int|Represents the number of categories or series between tick marks.|

### Customize axis labels

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|tickLabelPosition|string|Represents the position of tick-mark labels on the specified axis.|
|tickLabelSpacing|object|Represents the number of categories or series between tick-mark labels. Can be a value from 1 through 31999 or an empty string for automatic setting. The returned value is always a number.|


## Programmatic customize title and data label in chart

### Formatting substring of chart title
The `ChartTitle.getSubstring` method, which will return a [ChartFormatString](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/chartformatstring.md) object, is the entry point for programmatic format part of chart title. And `ChartFormatString.font` property which will return a [ChartFont](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/chartfont.md) object, is the formatting object which can be used to set font name, font size, color, etc.

### Customize chart title
The `ChartTitle` object have below new properties and methods to customize it

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|height|double|Returns the height, in points, of the chart title. Read-only. Null if chart title's not visible. Read-only.|
|horizontalAlignment|string|Represents the horizontal alignment for chart title. Possible values are: Center, Left, Right, Justify, Distributed.|
|left|double|Represents the distance, in points, from the left edge of chart title to the left edge of chart area. Null if chart title's not visible.|
|showShadow|bool|Represents a boolean value that determines if the chart title has a shadow.|
|textOrientation|int|Represents the text orientation of chart title. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|
|top|double|Represents the distance, in points, from the top edge of chart title to the top of chart area. Null if chart title's not visible.|
|verticalAlignment|string|Represents the vertical alignment of chart title. Possible values are: Center, Bottom, Top, Justify, Distributed.|
|width|double|Returns the width, in points, of the chart title. Read-only. Null if chart title's not visible. Read-only.|

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|setFormula(formula: string)|void|Sets a string value that represents the formula of chart title using A1-style notation.|

### Get & customize a datalabel from data point
The `ChartPoint.dataLabel` method, which will return a [ChartDataLabel](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/chartdatalabel.md) object, is the entry point for programmatic customize the datalabel.

The `ChartDataLabel` object have below new properties and methods to customize it

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|position|string|DataLabelPosition value that represents the position of the data label. Possible values are: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|
|separator|string|String representing the separator used for the data label on a chart.|
|showBubbleSize|bool|Boolean value representing if the data label bubble size is visible or not.|
|showCategoryName|bool|Boolean value representing if the data label category name is visible or not.|
|showLegendKey|bool|Boolean value representing if the data label legend key is visible or not.|
|showPercentage|bool|Boolean value representing if the data label percentage is visible or not.|
|showSeriesName|bool|Boolean value representing if the data label series name is visible or not.|
|showValue|bool|Boolean value representing if the data label value is visible or not.|

