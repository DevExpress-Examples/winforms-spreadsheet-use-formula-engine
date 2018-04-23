#region #usings
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;
using DevExpress.Spreadsheet.Functions;
#endregion #usings
using System;
using System.Globalization;
using System.Windows.Forms;

namespace FormulaEngineTest
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        Cell contextCell;
        Worksheet contextSheet;
        
        public Form1()
        {
            InitializeComponent();
            ribbonControl1.SelectedPage = formulaEngineRibbonPage1;
            spreadsheetControl1.Document.DocumentSettings.R1C1ReferenceStyle = true;
            editExpressionStyle.EditValue = 0;

            SetCellContext();
            spreadsheetControl1.DocumentLoaded += spreadsheetControl1_DocumentLoaded;
            spreadsheetControl1.ActiveCell.Formula = "SUM(R[1]C:R[2]C)";
            spreadsheetControl1.ActiveCell[1, 0].Formula = "ROW(R1:R10)";
            spreadsheetControl1.ActiveCell[2, 0].Formula = "ABS({-2,0,3})";
        }

        void spreadsheetControl1_DocumentLoaded(object sender, EventArgs e)
        {
            contextCell = spreadsheetControl1.ActiveCell;
            contextSheet = spreadsheetControl1.ActiveWorksheet;
        }

        private void btnSetContext_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            SetCellContext();
        }

        private void SetCellContext() {
            SetCellContext(spreadsheetControl1.ActiveCell);
        }
        private void SetCellContext(Cell activecell)
        {
            if (this.contextCell != null) contextCell.ClearFormats();
            this.contextCell = activecell;
            this.contextSheet = spreadsheetControl1.ActiveWorksheet;
            this.contextCell.FillColor = DevExpress.Utils.DXColor.SteelBlue;
        }

        private void btnR1C1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            spreadsheetControl1.Document.DocumentSettings.R1C1ReferenceStyle = !spreadsheetControl1.Document.DocumentSettings.R1C1ReferenceStyle;
        }


        private void btnEvaluate_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            #region #evaluate
            FormulaEngine engine = spreadsheetControl1.Document.FormulaEngine;
            // Get coordinates of the context cell.
            int columnIndex = this.contextCell.ColumnIndex;
            int rowIndex = this.contextCell.RowIndex;
            // Create the expression context.
            ExpressionContext context = new ExpressionContext(columnIndex, rowIndex, this.contextSheet,
                CultureInfo.GetCultureInfo("en-US"), ReferenceStyle.R1C1, DevExpress.Spreadsheet.Formulas.ExpressionStyle.Normal);
            // Evaluate the expression.
            ParameterValue result = engine.Evaluate("SUM(R[-2]C:R[-1]C)", context);
            // Get the result.
            double res = result.NumericValue;
            #endregion #evaluate
            MessageBox.Show(String.Format("The result is {0}", res),"Evaluate Test");
        }

        private void btnParse_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            #region #parse
            FormulaEngine engine = spreadsheetControl1.Document.FormulaEngine;
            if (spreadsheetControl1.ActiveCell != null){
                // Get the formula for parsing. 
                string formula = spreadsheetControl1.ActiveCell.Formula;
                if (formula != string.Empty)
                {
                    // Parse the formula
                    ParsedExpression p_expr = engine.Parse(formula);
                    // Traverse the expression tree and modify it.
                    MyVisitor visitor = new MyVisitor();
                    p_expr.Expression.Visit(visitor);
                    // Reconsitute the formula.
                    string formula1 = p_expr.ToString();
                    // Place the formula back to the cell.
                    spreadsheetControl1.ActiveCell.Formula = formula1;
                }
            #endregion #parse
            }
        }

        private void btnEvaluateExpressionStyle_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            #region #evaluateExpressionStyle
            FormulaEngine engine = spreadsheetControl1.Document.FormulaEngine;
            // Get coordinates of the active cell.
            int columnIndex = spreadsheetControl1.ActiveCell.ColumnIndex;
            int rowIndex = spreadsheetControl1.ActiveCell.RowIndex;
            // Create the expression context.
            ExpressionContext context = new ExpressionContext(columnIndex, rowIndex, this.contextSheet,
                CultureInfo.GetCultureInfo("en-US"), ReferenceStyle.UseDocumentSettings, (ExpressionStyle)editExpressionStyle.EditValue);
            // Evaluate the expression.
            ParameterValue result = engine.Evaluate(spreadsheetControl1.ActiveCell.Formula, context);
            // Get the result.
            string res = string.Empty;

            if (result.IsArray)
            {
                res += Environment.NewLine;
                int rowLength = result.ArrayValue.GetLength(0);
                int colLength = result.ArrayValue.GetLength(1);

                for (int i = 0; i < rowLength; i++)
                {
                    for (int j = 0; j < colLength; j++)
                    {
                        res += String.Format("{0} ", result.ArrayValue[i, j]);
                    }
                    res += Environment.NewLine;
                }
            }
            else res = result.ToString();
            #endregion #evaluateExpressionStyle
            MessageBox.Show(String.Format("The result is {0}", res), "Evaluate Using Expression Style");
        }
    }



}
