<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128613397/15.1.7%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T110344)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/FormulaEngineTest/Form1.cs) (VB: [Form1.vb](./VB/FormulaEngineTest/Form1.vb))
* [MyVisitor.cs](./CS/FormulaEngineTest/MyVisitor.cs) (VB: [MyVisitor.vb](./VB/FormulaEngineTest/MyVisitor.vb))
<!-- default file list end -->
# Formula Engine - examples of use


This example demonstrates how to use the FormulaEngine class for evaluating worksheet formulas and parsing expressions.<br>Select a cell and click the <strong>Set Context Cell</strong> button to specify a cell used to create the <a href="http://help.devexpress.com/#CoreLibraries/clsDevExpressSpreadsheetFormulasExpressionContexttopic">ExpressionContext  </a>object. The context cell is highlighted with blue color.<br>Click <strong>Evaluate Predefined Formula</strong> to calculate a simple formula <em>SUM(R[-2]C:R[-1]C)</em> by calling the <a href="http://help.devexpress.com/#CoreLibraries/DevExpressSpreadsheetFormulasFormulaEngine_Evaluatetopic">FormulaEngine.Evaluate</a> method with the specified context.<br>Click <strong>Parse and Modify Active Cell Formula</strong> to call the <a href="http://help.devexpress.com/#CoreLibraries/DevExpressSpreadsheetFormulasFormulaEngine_Parsetopic">FormulaEngine.Parse</a>  method and create an expression tree from the active cell formula. Subsequently a custom Visitor object is used to traverse the tree and increment row indexes which define an area referenced in the formula.<br>The <strong>Switch R1C1</strong> button can be used to change worksheet reference style if required to correctly interpret formula references.<br>To calculate formula contained in the active cell using different <a href="http://help.devexpress.com/#CoreLibraries/DevExpressSpreadsheetFormulasExpressionStyleEnumtopic">ExpressionStyle</a> settings, click <strong>Evaluate Cell Formula</strong>. When executed for the<em> ROW(R2:R11)</em> formula, the <strong>Normal</strong> expression style setting returns 1, the first row number, and the <strong>Array</strong> expression style returns an array of row numbers.

<br/>


