_[Workshop home](../../index.md)_  >  _[New Excel JavaScript APIs](../index.md)_ > _[Custom Properties API](index.md)_ > _Walkthrough_

# Walkthrough

In this first walkthrough we will open Script Lab and replace the code inside the Excel.run() with the following code:

```javascript
    let customDocProperties = context.workbook.properties.custom;
    let docProperties = context.workbook.properties;

    // update specific properties
    docProperties.comments = "Excel Demo was here!";

    // Add custom document properties.
    customDocProperties.add("ExcelDemo", "Property was set");

    await context.sync().then(() => {
        console.log("All done");
    });
```

Next, we run the script then click the Run button. Once complete we will see "All done!" in the console. You have successfully set the properties. Now that we have set the properties, lets go read them:

```javascript
    let customDocProperties = context.workbook.properties.custom;
    let docProperties = context.workbook.properties;
    let customProperty = customDocProperties.getItem("ExcelDemo");

    docProperties.load("comments");
    customProperty.load("key, value");

    await context.sync().then( () => {
        // get specific properties
        console.log("Comments: " + docProperties.comments);

        // Get custom document properties.
        console.log("ExcelDemo: " + customProperty.value);
    });
```

This point we will see this:
![alt text](images/walkthrough.PNG?raw=true "walkthrough results")
