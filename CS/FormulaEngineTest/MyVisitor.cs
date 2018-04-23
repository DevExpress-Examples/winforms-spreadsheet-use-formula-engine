namespace FormulaEngineTest
{
    #region #myvisitor
    public class MyVisitor : DevExpress.Spreadsheet.Formulas.ExpressionVisitor
    {
        public override void Visit(DevExpress.Spreadsheet.Formulas.CellReferenceExpression expression)
        {
            base.Visit(expression);

            expression.CellArea.TopRowIndex++;
            expression.CellArea.BottomRowIndex++;
        }
    }
    #endregion #myvisitor
}
