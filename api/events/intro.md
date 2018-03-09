_[Workshop home](../../index.md)_  >  _[New Excel JavaScript APIs](../index.md)_ > _[Event API additions](index.md)_ > _Introduction_

# Event API introduction

## Event API in Beta
Events APIs in JavaScript provides a way of interacting between add-ins and users upon several objects. Each time certain types of changes occur in Excel, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific action occurs. 

The actions include operations such as changing selection on a worksheet, adding or deleting a worksheet, changing the contents in a cell, etc. Actions can come from both local and remote users, who are co-authoring on the same workbook. Developers can define a JavaScript function called “Event handler” in their add-in and register it to a specific event in Excel. When the events occur in the workbook, the event handler will be run automatically. In the event handler, developers can call other Excel JavaScript APIs to further interact with the Workbook, with the information carried in the event arguments.

The following events are currently supported.

**Below are the new events:**

| Object | Event | Description | Event Argument |
| --- | --- | --- | --- |
| Table | onChanged | Occurs when cells on the table are changed by the user or by APIs. | [TableChangedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/tablechangedeventargs.md) |
| Table | onSelectionChanged | Occurs when selection has changed on a table. | [TableSelectionChangedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/tableselectionchangedeventargs.md) |
| TableCollection | onChanged | Occurs when cells on any table are changed by the user or by APIs. | [TableChangedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/tablechangedeventargs.md) |
| Worksheet | onActivated | Occurs when the worksheet has become activated. | [WorksheetActivatedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetactivatedeventargs.md) |
| Worksheet| onChanged | Occurs when cells on the worksheet are changed by the user or APIs. | [WorksheetChangedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetchangedeventargs.md) |
| Worksheet | onDeactivated | Occurs when the worksheet has become deactivated. | [WorksheetDeactivatedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetdeactivatedeventargs.md) |
| Worksheet | onSelectionChanged | Occurs when selection has changed on the worksheet | [WorksheetSelectionChangedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetselectionchangedeventargs.md) |
| WorksheetCollection | onActivated | Occurs when any worksheet has become activated. | [WorksheetActivatedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetactivatedeventargs.md) |
| WorksheetCollection | onAdded | Occurs when a worksheet has been added to the workbook. | [WorksheetAddedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetaddedeventargs.md) |
| WorksheetCollection| onDeactivated | Occurs when any worksheet has become deactivated. | [WorksheetDeactivatedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetdeactivatedeventargs.md) |
| WorksheetCollection | onDeleted | Occurs when a worksheet has been deleted from the workbook. | [WorksheetDeletedEventargs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetdeletedeventargs.md) |

**More events will be upcoming soon:**

| Object | Event | Description | Event Argument |
| --- | --- | --- | --- |
| Chart | onActivated | Occurs when the chart has become activated. | [ChartActivatedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/chartactivatedeventargs.md) |
| Chart | onDeactivated | Occurs when the chart has become deactivated. | [ChartDeactivatedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/chartdeactivatedeventargs.md) |
| ChartCollection | onActivated | Occurs when any chart has become activated. | [ChartActivatedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/chartactivatedeventargs.md) |
| ChartCollection | onAdded | Occurs when a chart has been added. | [ChartAddedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/chartaddedeventargs.md) |
| ChartCollection | onDeactivated | Occurs when any chart has become deactivated | [ChartDeactivatedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/chartdeactivatedeventargs.md) |
| ChartCollection | onDeleted | Occurs when a worksheet has been deleted from the workbook. | [ChartDeletedEvent](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/chartdeletedevent.md) |
| Worksheet | onCalculated | Occurs when the workbook has finished calculation. | [WorsheetCalculatedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetcalculatedeventargs.md) |
| WorkbookCollection | onCalculated | Occurs when all the worksheets of the workbook have finished calculation. | [WorsheetCalculatedEventArgs](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetcalculatedeventargs.md) |

**Other new APIs related:**

1. Added **eventArgs.getRange()** method for the **onChanged** event to return the range object that is associated with the address when the event occurs.
2. Added **context.runtime.enableEvents** (true/false) to turn JavaScript events on and off for the current taskpane or content add-in.
3. Added **application.calculationMode** to change the calculation mode of Excel. Options are _Automatic, AutomaticExcepTables, Manual_.
