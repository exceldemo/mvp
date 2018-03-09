_[Workshop home](../../index.md)_  >  _[New Excel JavaScript APIs](../index.md)_ > _[Data Validation API](index.md)_ > _Introduction_

# Data Validation

## Add data validation to Excel ranges

The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook. To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:

- [Apply data validation to cells](https://support.office.com/en-us/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [More on data validation](https://microsoft.sharepoint.com/:p:/r/teams/oext/_layouts/15/Doc.aspx?sourcedoc=%7B51143964-d52c-429d-bfac-c7495473d536%7D&action=edit)
- [Description and examples of data validation in Excel](https://support.microsoft.com/en-us/help/211485/description-and-examples-of-data-validation-in-excel)

## Programmatic control of data validation

The `Range.dataValidation` property, which takes a  [DataValidation](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/datavalidation.md) object, is the entry point for programmatic control of data validaiton in Excel. There are five properties to the `DataValidation` object:

- `rule` &#8212; Defines what constitutes valid data for the range. See [DataValidationRule](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/datavalidation.md).
- `errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**. See [DataValidationErrorAlert](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/datavalidationerroralert.md).
- `prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message. See [DataValidationPrompt](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/datavalidationrule.md).
- `ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range. Defaults to `true`.
- `type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.

> [!NOTE]
> Data validation added programmatically behaves just like manually added data validation. In particular, note that data validation is triggered only if the user directly types a value into a cell or copies and pastes a cell from elsewhere in the workbook and takes the **Values** paste option. If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.

### Creating validation rules

To add data validation to a range, your code must set the `rule` proeprty of the `DataValidation` object in `Range.dataValidation`. This takes a [DataValidationRule](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/datavalidationrule.md) object which has seven optional properties. *No more than one of these properties may be present in any `DataValidationRule` object.* The property that you include determines the type of validation.

