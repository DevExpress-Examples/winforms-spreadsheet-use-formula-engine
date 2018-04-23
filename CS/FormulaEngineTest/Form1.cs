#region #usings
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;
using DevExpress.Spreadsheet.Functions;
#endregion #usings
using DevExpress.Utils;
using System;
using System.Globalization;
using System.Windows.Forms;

namespace FormulaEngineTest
{
    public partial class Form1 : Form
    {
        Cell contextCell;
        Worksheet contextSheet;
        
        public Form1()
        {
            InitializeComponent();
            ribbonControl1.SelectedPage = formulaEngineRibbonPage1;
            spreadsheetControl1.Document.DocumentSettings.R1C1ReferenceStyle = true;
            SetCellContext();
            spreadsheetControl1.ActiveCell.Formula = "SUM(R[1]C:R[2]C)";

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
            this.contextCell.FillColor = DXColor.SteelBlue;
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
                CultureInfo.GetCultureInfo("en-US"),false, ReferenceStyle.R1C1);
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

    }



}
