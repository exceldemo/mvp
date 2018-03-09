_[Workshop home](../../index.md)_  >  _[New Excel JavaScript APIs](../index.md)_ > _[Chart API additions](index.md)_ > _Walkthrough_


# Walkthrough

Here are some examples to programatic control series/trendline/title/axis in a chart. You can try below scripts using ScriptLab in this [Sample file](sampleDoc/ExcelChartAPISample.xlsx) and import this [gist](https://gist.github.com/binwang2017/cd94945f613323205393bd7c9f79a552)

## Add/delete/customize series

### Add a series

```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let valueRange = sheet.getRange("B1:B10");
    let categoryRange = sheet.getRange("A1:A10");
    let series = chart.series.add();
    series.setXAxisValues(categoryRange);
    series.setValues(valueRange);

    return context.sync();
})
```
Result

![workthrough_1](image/workthrough_1.PNG?raw=true)

### Customize a series

In this example, we customized the new added series's marker style to circle, background color to <font color='red'>red</font> and make all datalabels belong to this series visible. 

```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    series.markerStyle = Excel.ChartMarkerStyle.circle;
    series.markerBackgroundColor = 'red';
    series.hasDataLabels = true;

    return context.sync();
})
```

Result

![workthrough_2](image/workthrough_2.PNG?raw=true)

### Delete a series

```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    series.delete();

    return context.sync();
})
```

Result

![workthrough_3](image/workthrough_3.PNG?raw=true)

## Add/delete/customize trendline

### Add a trendline

Before add a trendline, we need add a series to chart follow the example [Add a series]() above

```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    let trendline = series.trendlines.add(Excel.TrendlineType.polynomial);

    return context.sync();
})
```

Result

![workthrough_4](image/workthrough_4.PNG?raw=true)

### Customize a trendline

In this example, we customized the new added trendline to display the euqation and make the line color to <font color='green'>green</font> and style to dash-dot-dot.

```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    let trendline = series.trendlines.getItem(0);
    trendline.displayEquation = true;
    trendline.format.line.color = 'green'
    trendline.format.line.lineStyle = Excel.ChartLineStyle.dashDotDot;

    return context.sync();
})
```

Result

![workthrough_5](image/workthrough_5.PNG?raw=true)

### Delete a trendline

```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    let trendline = series.trendlines.getItem(0);
    trendline.delete();

    return context.sync();
})
```

Result

![workthrough_1](image/workthrough_1.PNG?raw=true)

## Customize chart title and datalabel

### Set font format of part of chart title

In this example, we customized the chart title to <b>This is a chart</b> and highlight the <font color='red'>chart</font>.

```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    chart.title.text = "This is a chart";
    let textrange = chart.title.getSubstring(10, 5);
    let font = textrange.font;
    font.size = 18;
    font.color = 'red';
    font.bold = true;
    font.italic = true;

    return context.sync();
})
```

Result

![workthrough_6](image/workthrough_6.PNG?raw=true)

### Customize datalabel & point

In this example, we first make the 6th datalabel visible and then customized it to show legend key and category name and put its position to the point's left. And we customized corresponding point's marker size to 8, background color to <font color='orange'>orange</font> and the marker style to diamond. 

```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    let point = series.points.getItemAt(5); // 6th datalabel
    point.hasDataLabel = true;
    point.markerSize = 8;
    point.markerBackgroundColor = 'orange';
    point.markerStyle = Excel.ChartMarkerStyle.diamond;
    let datalabel = point.dataLabel;
    datalabel.showCategoryName = true;
    datalabel.showValue = true;
    datalabel.showLegendKey = true;
    datalabel.position = Excel.ChartDataLabelPosition.left;

    return context.sync();
})
```

Result

![workthrough_7](image/workthrough_7.PNG?raw=true)

## Customize chart axis

In this example, we customized the value axis display unit to based on 20 and the category axis's tick label spacing to 2.

```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let valueAxis = chart.axes.valueAxis;
    valueAxis.setCustomDisplayUnit(20);
    let categoryAxis = chart.axes.categoryAxis;
    categoryAxis.tickLabelSpacing = 2;

    return context.sync();
})
```
Result

![workthrough_8](image/workthrough_8.PNG?raw=true)





 
