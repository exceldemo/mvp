_[Workshop home](../../index.md)_  >  _[New Excel JavaScript APIs](../index.md)_ > _[Event API additions](index.md)_ > _Walkthrough_

# Walkthrough 

## Event triggers 

Events within Excel can be triggered by:

User interaction via the Excel user interface (UI)
Office add-in (JavaScript) code
VBA add-in (macro) code
Any change that complies with default behavior of Excel will trigger the corresponding event(s).

## Lifecycle of an event handler 
An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.

## Events and coauthoring
With coauthoring, multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as onChanged, the corresponding Event object will contain a source property that indicates whether the event was triggered locally by the current user (event.source = Local) or was triggered by the remote coauthor (event.source = Remote).

## How to Use 

Below are the steps to use the JavaScript events. Please insert **Script Lab** add-in and try the events.

You can [import simple sample script](https://gist.github.com/1f943822437ec35d2ea5c9b3a0efc138) into Script Lab and try it by yourself.

Step1: Define event handler

```js
async function eventHandler(eventArgs) {
    await Excel.run(async (context) => {
        var worksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId);
        worksheet.load();
        awaitcontext.sync();
        console.log("Event received on " + worksheet.name);
        console.log(JSON.stringify(eventArgs));
    })
}
```


Step2: Register the events

```js
var eventResult;
async function registerWorksheetEvents() {
    await Excel.run(async (context) => {
        var worksheet = context.workbook.worksheets.getActiveWorksheet();
        eventResult = worksheet.onSelectionChanged.add(eventHandler);
        worksheet.load();
        await context.sync();
        console.log("Events registered successfully on " + worksheet.name);
    });
}
```


Step3: Observe the events

Make some selection change on the active worksheet on which the event has been registered. Check the logs in the Script Lab console and verify the event details.


![output](images/Picture1.png?raw=true)

Step4: Unregister the events

```js
async function unregisterWorksheetEvents() {
    await Excel.run (eventResult.context, async (context) => {
        if (eventResult) {
            eventResult.remove();
        }
        await context.sync();
        console.log("Events unregistered successfully.");
    });
}
```

**Other Reference Samples:**

Events Driven Bing Maps: [https://gist.github.com/e3446002107f06c9ac450e6519e4f793](https://gist.github.com/e3446002107f06c9ac450e6519e4f793)

If you have any feedback or question, please feel free to contact the feature crew: [ecoxleventsapi@microsoft.com](mailto:ecoxleventsapi@microsoft.com)
