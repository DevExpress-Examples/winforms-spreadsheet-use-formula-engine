Imports Microsoft.VisualBasic
Imports System
Namespace FormulaEngineTest
	#Region "#myvisitor"
	Public Class MyVisitor
		Inherits DevExpress.Spreadsheet.Formulas.ExpressionVisitor
		Public Overrides Sub Visit(ByVal expression As DevExpress.Spreadsheet.Formulas.CellReferenceExpression)
			MyBase.Visit(expression)

			expression.CellArea.TopRowIndex += 1
			expression.CellArea.BottomRowIndex += 1
		End Sub
	End Class
	#End Region ' #myvisitor
End Namespace
