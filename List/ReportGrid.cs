namespace GBTH.List
{
    public class ReportGrid : DataGridView
    {
        public ReportGrid() : base()
        {
            this.Columns.Add("Number", "번호"); //0
            this.Columns.Add("Date", "날짜");
            this.Columns.Add("Descrytion", "내용");
            this.Columns.Add("Insert", "입고"); //3
            this.Columns.Add("Use", "사용"); //4
            this.Columns.Add("Storege", "재고"); //5
            this.Columns.Add("Manager", "담당자");
            this.Columns.Add("Other", "비고");

            this.Columns[0].ValueType = typeof(int);
            this.Columns[3].ValueType = typeof(int);
            this.Columns[4].ValueType = typeof(int);
            this.Columns[5].ValueType = typeof(int);
            this.Columns[5].ReadOnly = true;

            this.Click += ReportDataSetting;
            this.CellValueChanged += ReportDataSetting;
            this.RowsAdded += ReportGridRowsAdded;
            this.MultiSelect = false;

            foreach (DataGridViewColumn column in this.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void ReportGridRowsAdded(object? sender, DataGridViewRowsAddedEventArgs e)
        {
            for (int r = 0; r < this.RowCount - 1; r++)
            {
                this.Rows[r].Cells[0].Value = r + 1;
            }
        }

        private void ReportDataSetting(object? sender, EventArgs e)
        {
            this.Rows[0].Cells[5].Value
                    = int.Parse((this.Rows[0].Cells[3].EditedFormattedValue as string == "" ? "0"
                        : this.Rows[0].Cells[3].EditedFormattedValue as string)!);

            for (int r = 1; r < this.RowCount; r++)
            {
                this.Rows[r].Cells[5].Value
                    = int.Parse((this.Rows[r - 1].Cells[5].EditedFormattedValue as string == "" ? "0"
                        : this.Rows[r - 1].Cells[5].EditedFormattedValue as string)!)
                    - int.Parse((this.Rows[r].Cells[4].EditedFormattedValue as string == "" ? "0"
                        : this.Rows[r].Cells[4].EditedFormattedValue as string)!)
                    + int.Parse((this.Rows[r].Cells[3].EditedFormattedValue as string == "" ? "0"
                        : this.Rows[r].Cells[3].EditedFormattedValue as string)!);

            }
        }
    }
}
