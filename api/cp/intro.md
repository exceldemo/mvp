_[Workshop home](../../index.md)_  >  _[New Excel JavaScript APIs](../index.md)_ > _[Custom Properties API](index.md)_ > _Introduction_

# Custom Properties

## Document Properties API 

This covers:
* Document Properties 
* Custom Document Proeprties

The document properties allows you to read and write both the built-in document properties as well as add/remove custom document properties. The following are the supported methods and proeprties:

### context.workbook.properties
The following are the supported properties:

| Property Name | Description | Type |
|---------------|-------------|------|
| custom  | Custom Document Properties  | CustomDocumentProperties |
| author | The author of the document | (string) |
| lastAuthor | The last person to modify the workbook | (string) |
| revisionNumber | The current revision of the workbook | (string) |
| title | The title of the workbook  | (string) |
| subject | The subject of the workbook | (string) |
| keywords | Custom keywords that can be added to help identify/find the workbook | (string) |
| comments | Comments about the workbook | (string) |
| category | Workbook category | (string) |
| manager | The name of a manager | (string) |
| company | A company name | (string) |
| creationDate | Date and time the document was created | (DateTime) |

The following are the supported methods:

| Method | Name | Description |
|--------|------|-------------|
| load | Loads the item after a context.sync() | Object. load("author, lastAuthor, revisionNumber, title, subject, keywords, comments, category, manager, company, creationDate") |

## Custom Document Properties API
The following are supported methods of the custom document properties:

### context.workbook.properties.customdocumentproperties

| Methods | Description |   |
|---------|-------------|---|
| getItem | The name of the custom document property to get | Object.getItem(“name”) **Returns** CustomDocumentProperty object |
| add | Adds/Updates a given property on the next context.sync() | Object.add(“key”,”value”) |


### context.workbook.properties.customdocumentproperty
The following are the supported methods for the customdocumentproeprty object returned from the getItem call.

| Method Name | Description | Type |
|---------------|-------------|------|
| load | Loads the item after a context.sync() | Object.load(“key”,”value”) |
| delete | Deletes the custom property | (n/a) |

The following are the supported properties for the customdocumentproeprty object returned from the getItem call

| Property Name | Description | Type |
|---------------|-------------|------|
| key | The key for the loaded property | (string) |
| Value | The value for the loaded property | (string) |

