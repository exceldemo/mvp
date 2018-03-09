_[Workshop home](../../index.md)_  >  _[New Excel JavaScript APIs](../index.md)_ > _[Data Validation API](index.md)_ > _Walkthrough_

# Walkthrough

## Basic and DateTime validation rule types

The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/basicdatavalidation.md) object as their value.

- `wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.
- `decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.
- `textLength` &#8212; Applies the validation details in the the `BasicDataValidation` object to the *length* of the cell's value.

Here is an example of creating a validation rule. Note the following about this code:

- The `operator` is the binary operator "GreaterThan". Whenever you use a binary operator, the value that the user trys to enter in the cell is the left hand operand and the value specified in `formula1` is the right hand operand. So this rule says that only whole numbers that are greater than 0 are valid. 
- The `formula1` is a hard-coded number. If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value. For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            wholeNumber: {
                formula1: 0,
                operator: "GreaterThan"
            }
        };

    return context.sync();
})
```

See [BasicDataValidation](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/basicdatavalidation.md) for a list of the other binary operators. 

There are also two ternary operators: "Between" and "NotBetween". To use these, you must specify the optional `formula2` property. The `formula1` and `formula2` values are the bounding operands. The value that the user trys to enter in the cell is the third (evaluated) operand. The following is an example of using the "Between" operator:

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            decimal: {
                formula1: 0,
                formula2: 100,
                operator: "Between"
            }
        };

    return context.sync();
})
```

The next two rule properties take a [DateTimeDataValidation](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/basicdatavalidation.md) object as their value.

- `date`
- `time`

The `DateTimeDataValidation` object is structured just like the `BasicDataValidation`, with properties `formula1`, `formula2`, and `operator`, and is used in the same way. The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula). The following is an example that defines valid values as dates in the first week of April, 2018. 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            date: {
                formula1: "2018-04-01",
                formula2: "2018-04-08",
                operator: "Between"
            }
        };

    return context.sync();
})
```

### The List validation rule type

Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list. The following is an example. Note the following about this code:

- It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.
- The `source` property specifies the list of valid values. The range with the names has been assigned to it. You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz". 
- The `inCellDropDown` property specifies whether a drop down control will appear in the cell when the user selects it. If set to true, then the drop down appears with the list of values from the `source`.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: nameSourceRange
        }
    };

    return context.sync();
})
```

### The Custom validation rule type

Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula. The following is an example. Note the following about this code:

- It assumes there is a two column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.
- To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.
- `SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2. If A2 is not contained in B2, it does not return a number. `ISNUMBER()` returns a boolean. So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var commentsRange = sheet.tables.getItem("AtheletesTable").columns.getItem("Comments").getDataBodyRange();

    commentsRange.dataValidation.rule = {
            custom: {
                formula: "=NOT(ISNUMBER(SEARCH(A2,B2)))"
            }
        };

    return context.sync();
})
```

## Creating validation error alerts

You can create custom error alerts that will appear when a user tries to enter invalid data in a cell. The following is a simple example. Note the following about this code:

- The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert. Only `Stop` will actually prevent the user from adding invalid data. The popup for `Warning` and `Information` has options that let the user enter the invalid data anyway.
- The `showAlert` property defaults to `true`. This means that the Office host will popup a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.


```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.errorAlert = {
            message: "Sorry, only positive whole numbers are allowed",
            showAlert: true, // default is 'true'
            style: "Stop", // other possible values: Warning, Information
            title: "Negative or Decimal Number Entered"
        };
    
    // Set range.dataValidation.rule and optionally .prompt here.

    return context.sync();
})
```

For more information, see [DataValidationErrorAlert](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/datavalidationerroralert.md).

## Creating validation prompts

You can create an instructional prompt that will appear when user's hover over, or select, a cell to which data validation has been applied. The following is an example:

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.prompt = {
            message: "Please enter a positive whole number.",
            showPrompt: true, // default is 'false'
            title: "Positive Whole Numbers Only."
        };
    
    // Set range.dataValidation.rule and optionally .errorAlert here.

    return context.sync();
})
```

For more information, see [DataValidationPrompt](hhttps://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/datavalidationprompt.md).

## Remove data validation from a range

To remove data validation from a range, call the  [Range.dataValidation.clear()](https://dev.office.com/reference/add-ins/excel/datavalidation#clear) method.

```js
myrange.dataValidation.clear()
```

It is not necessary that the range you clear is exactly the same range as a range on which you added data validation. If it isn't, only the overlapping cells, if any, of the two ranges are cleared. 

> [!NOTE]
> Clearing data validation from a range will also clear any data validation that a user has added manually to the range.


 
