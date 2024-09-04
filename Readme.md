<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128613397/24.2.1%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T110344)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->

# Spreadsheet for WinForms - Use Formula Engine

This example demonstrates how to use the [FormulaEngine](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Formulas.FormulaEngine) class to evaluate worksheet formulas and parse expressions.

## Implementation Details

Select a cell and click the **Set Context Cell** button to specify a cell used to create the [ExpressionContext](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Formulas.ExpressionContext) object. The context cell is highlighted with blue color.

Click **Evaluate Predefined Formula** to calculate a simple formula _SUM(R[-2]C:R[-1]C)_ by the the [FormulaEngine.Evaluate](https://docs.devexpress.com/OfficeFileAPI/devexpress.spreadsheet.formulas.formulaengine.evaluate.overloads) method call with the specified context.

Click **Parse and Modify Active Cell Formula** to call the [FormulaEngine.Parse](https://docs.devexpress.com/OfficeFileAPI/devexpress.spreadsheet.formulas.formulaengine.parse.overloads) method and create an expression tree from the active cell formula. Subsequently a custom `Visitor`` object is used to traverse the tree and increment row indexes which define an area referenced in the formula.

The **Switch R1C1** button can be used to change worksheet reference style if required to correctly interpret formula references.

To calculate formula contained in the active cell using different [ExpressionStyle](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Formulas.ExpressionStyle) settings, click **Evaluate Cell Formula**. When executed for the _ROW(R2:R11)_ formula, the **Normal** expression style setting returns **1**, the first row number, and the `Array` expression style returns an array of row numbers.

## Files to Review

* [Form1.cs](./CS/FormulaEngineTest/Form1.cs) (VB: [Form1.vb](./VB/FormulaEngineTest/Form1.vb))
* [MyVisitor.cs](./CS/FormulaEngineTest/MyVisitor.cs) (VB: [MyVisitor.vb](./VB/FormulaEngineTest/MyVisitor.vb))

## Documentation

* [Formula Engine](https://docs.devexpress.com/WindowsForms/17143/controls-and-libraries/spreadsheet/formulas/formula-engine)
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=winforms-spreadsheet-use-formula-engine&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=winforms-spreadsheet-use-formula-engine&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
