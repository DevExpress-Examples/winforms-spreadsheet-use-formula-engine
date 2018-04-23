Imports Microsoft.VisualBasic
#Region "#usings"
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas
Imports DevExpress.Spreadsheet.Functions
#End Region ' #usings
Imports DevExpress.Utils
Imports System
Imports System.Globalization
Imports System.Windows.Forms

Namespace FormulaEngineTest
	Partial Public Class Form1
		Inherits Form
		Private contextCell As Cell
		Private contextSheet As Worksheet

		Public Sub New()
			InitializeComponent()
			ribbonControl1.SelectedPage = formulaEngineRibbonPage1
			spreadsheetControl1.Document.DocumentSettings.R1C1ReferenceStyle = True
			SetCellContext()
			spreadsheetControl1.ActiveCell.Formula = "SUM(R[1]C:R[2]C)"

		End Sub

		Private Sub btnSetContext_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSetContext.ItemClick
			SetCellContext()
		End Sub

		Private Sub SetCellContext()
			SetCellContext(spreadsheetControl1.ActiveCell)
		End Sub
		Private Sub SetCellContext(ByVal activecell As Cell)
			If Me.contextCell IsNot Nothing Then
				contextCell.ClearFormats()
			End If
			Me.contextCell = activecell
			Me.contextSheet = spreadsheetControl1.ActiveWorksheet
			Me.contextCell.FillColor = DXColor.SteelBlue
		End Sub

		Private Sub btnR1C1_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnR1C1.ItemClick
			spreadsheetControl1.Document.DocumentSettings.R1C1ReferenceStyle = Not spreadsheetControl1.Document.DocumentSettings.R1C1ReferenceStyle
		End Sub


		Private Sub btnEvaluate_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnEvaluate.ItemClick
'			#Region "#evaluate"
			Dim engine As FormulaEngine = spreadsheetControl1.Document.FormulaEngine
			' Get coordinates of the context cell.
			Dim columnIndex As Integer = Me.contextCell.ColumnIndex
			Dim rowIndex As Integer = Me.contextCell.RowIndex
			' Create the expression context.
            Dim context As New ExpressionContext(columnIndex, rowIndex, Me.contextSheet, _
                                                 CultureInfo.GetCultureInfo("en-US"), False, ReferenceStyle.R1C1)
			' Evaluate the expression.
			Dim result As ParameterValue = engine.Evaluate("SUM(R[-2]C:R[-1]C)", context)
			' Get the result.
			Dim res As Double = result.NumericValue
'			#End Region ' #evaluate
			MessageBox.Show(String.Format("The result is {0}", res),"Evaluate Test")
		End Sub

		Private Sub btnParse_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnParse.ItemClick
'			#Region "#parse"
			Dim engine As FormulaEngine = spreadsheetControl1.Document.FormulaEngine
			If spreadsheetControl1.ActiveCell IsNot Nothing Then
				' Get the formula for parsing. 
				Dim formula As String = spreadsheetControl1.ActiveCell.Formula
				If formula <> String.Empty Then
					' Parse the formula
					Dim p_expr As ParsedExpression = engine.Parse(formula)
					' Traverse the expression tree and modify it.
					Dim visitor As New MyVisitor()
					p_expr.Expression.Visit(visitor)
					' Reconsitute the formula.
					Dim formula1 As String = p_expr.ToString()
					' Place the formula back to the cell.
					spreadsheetControl1.ActiveCell.Formula = formula1
				End If
'			#End Region ' #parse
			End If
		End Sub

	End Class



End Namespace
