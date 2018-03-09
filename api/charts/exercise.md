_[Workshop home](../../index.md)_  >  _[New Excel JavaScript APIs](../index.md)_ > _[Chart API additions](index.md)_ > _Exercise_

# Exercise

Here is a table(SalesTable) about the Sales of Lemon and Orange in July as below.

![Data](image/data.PNG?raw=true)

Open the sample [document](sampleDoc/ExcelChartAPISample.xlsx) and navigate to <b>Exercise</b> tab. And using Script Lab, import this [gist](https://gist.github.com/binwang2017/4efe892bbe2fca5e430f3c3593ac3f64) then start your exercise.

#### Step 1 

And in this section please create a new series in the exsiting empty chart that will show total sales(*column H*) in July and add a trendline(*Excel.TrendlineType.polynomial*) show as below

![Step 1 Result](image/Step_1_Result.PNG?raw=true)

Some usefull code snippets for exercise

- Get Exercise sheet & chat in sample doc

```js
let sheet = context.workbook.worksheets.getItem("Exercise");
let chart = sheet.charts.getItemAt(0);
```

- Get date and sales range in table

```js
let table = context.workbook.tables.getItem("SalesTable");
let salesRange = table.columns.getItem("Total Sales").getDataBodyRange();
let dateRange = table.columns.getItem("Date").getDataBodyRange();
```

- Another way to set display unit for value axis

```js
let valueAxis = chart.axes.valueAxis;
valueAxis.displayUnit = Excel.ChartAxisDisplayUnit.thousands;
```

#### Step 2
Highlight the highest sales(the 27th data point) in the series as below
![Step 2 Result](image/Step_2_Result.PNG?raw=true)


